import { readFileSync } from 'fs';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { z } from 'zod';
import type { GraphClient } from '../src/graph-client.js';
import {
  buildToolsRegistry,
  registerGraphTools,
  type GraphToolRegistrationOptions,
} from '../src/graph-tools.js';
import { getCombinedPresetPattern, TOOL_CATEGORIES } from '../src/tool-categories.js';
import { describeToolSchema } from '../src/lib/tool-schema.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn(), debug: vi.fn() },
}));
vi.mock('../src/audit-log.js', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/audit-log.js')>()),
  auditLog: vi.fn(),
}));

// Exercise the actual generated client and endpoint config, not mocked endpoint schemas.
const contracts = [
  [
    'accept-shared-calendar-event',
    'POST',
    'events/{event-id}/accept',
    'Calendars.ReadWrite.Shared',
    false,
  ],
  [
    'tentatively-accept-shared-calendar-event',
    'POST',
    'events/{event-id}/tentativelyAccept',
    'Calendars.ReadWrite.Shared',
    false,
  ],
  [
    'decline-shared-calendar-event',
    'POST',
    'events/{event-id}/decline',
    'Calendars.ReadWrite.Shared',
    false,
  ],
  ['get-shared-calendar-event', 'GET', 'events/{event-id}', 'Calendars.Read.Shared', true],
  [
    'list-shared-calendar-event-instances',
    'GET',
    'events/{event-id}/instances',
    'Calendars.Read.Shared',
    true,
  ],
  ['find-shared-meeting-times', 'POST', 'findMeetingTimes', 'Calendars.Read.Shared', true],
  ['list-shared-calendars', 'GET', 'calendars', 'Calendars.Read.Shared', true],
  [
    'get-shared-specific-calendar-view',
    'GET',
    'calendars/{calendar-id}/calendarView',
    'Calendars.Read.Shared',
    true,
  ],
  ['list-shared-calendar-view-delta', 'GET', 'calendarView/delta()', 'Calendars.Read.Shared', true],
  [
    'list-shared-mailbox-folder-messages-delta',
    'GET',
    'mailFolders/{mailFolder-id}/messages/delta()',
    'Mail.Read.Shared',
    true,
  ],
  [
    'list-shared-calendar-event-attachments',
    'GET',
    'events/{event-id}/attachments',
    'Calendars.Read.Shared',
    true,
  ],
  [
    'get-shared-calendar-event-attachment-content',
    'GET',
    'events/{event-id}/attachments/{attachment-id}/$value',
    'Calendars.Read.Shared',
    true,
  ],
] as const;

const allowedScopes = 'Calendars.Read.Shared Calendars.ReadWrite.Shared Mail.Read.Shared';
const owner = 'executive+calendar@contoso.org';
const eventId = 'occurrence/+id=';
const eventPath = '/users/executive%2Bcalendar%40contoso.org/events/occurrence%2F%2Bid=';
const windowParams = { startDateTime: '2026-12-01T00:00:00Z', endDateTime: '2026-12-08T00:00:00Z' };
const jsonResult = (data: unknown) => ({ content: [{ type: 'text', text: JSON.stringify(data) }] });

function setup(options: GraphToolRegistrationOptions = {}) {
  const server = { tool: vi.fn(), registerTool: vi.fn(), server: { _requestHandlers: new Map() } };
  const graphRequest = vi.fn().mockResolvedValue(jsonResult({ value: [] }));
  const client = { graphRequest, serialize: JSON.stringify } as unknown as GraphClient;
  registerGraphTools(server, client, {
    orgMode: true,
    enabledTools: getCombinedPresetPattern(['work']),
    allowedScopes,
    ...options,
  });
  const tools = new Map(
    server.registerTool.mock.calls.map(([name, config, handler]) => [name, { config, handler }])
  );
  const call = async (name: string, params: Record<string, unknown>) => {
    const registered = tools.get(name);
    expect(registered, `registered ${name}`).toBeDefined();
    // The real MCP server validates against this schema before invoking the handler.
    return registered!.handler(registered!.config.inputSchema.parse(params));
  };
  return { tools, graphRequest, call };
}

describe('delegated calendar tools', () => {
  beforeEach(() => {
    vi.stubEnv('MS365_MCP_MAX_PAGES', '5');
    vi.stubEnv('MS365_MCP_BODY_FORMAT', 'text');
    vi.stubEnv('MS365_MCP_REQUIRE_CONFIRM', 'true');
  });
  afterEach(() => vi.unstubAllEnvs());

  it.each(contracts)(
    '%s registers with the reviewed path, scopes and approval annotation',
    (name, method, path, scope, readOnly) => {
      const { tools } = setup();
      const entry = buildToolsRegistry(false, true).get(name)!;
      expect(entry.config).toMatchObject({
        pathPattern: `/users/{user-id}/${path}`,
        method: method.toLowerCase(),
        workScopes: [scope],
        presets: ['work'],
      });
      expect(entry.config?.scopes).toBeUndefined();
      expect(tools.get(name)?.config.annotations).toMatchObject({
        readOnlyHint: readOnly,
        destructiveHint: !readOnly,
      });
      expect(tools.get(name)?.config.description).toContain('ID or UPN');
      expect(tools.get(name)?.config.description).toContain('delegated access');
      for (const [preset, category] of Object.entries(TOOL_CATEGORIES)) {
        if (preset === 'all') continue;
        expect(new RegExp(category.pattern.source, 'i').test(name), preset).toBe(preset === 'work');
      }
    }
  );

  it('excludes every new tool in personal mode and preserves reads in read-only mode', () => {
    const personal = setup({ orgMode: false });
    const readonly = setup({ readOnly: true });
    const discovery = buildToolsRegistry(true, true);
    for (const [name, , , , readOnly] of contracts) {
      expect(personal.tools.has(name)).toBe(false);
      expect(readonly.tools.has(name)).toBe(readOnly);
      expect(discovery.has(name)).toBe(readOnly);
    }
  });

  it('does not expose shared writes with read scopes, or findMeetingTimes in app mode', () => {
    const reads = setup({ allowedScopes: 'Calendars.Read.Shared Mail.Read.Shared' });
    for (const [name, , , , readOnly] of contracts) expect(reads.tools.has(name)).toBe(readOnly);
    expect(
      setup({ appMode: true, allowedScopes: undefined }).tools.has('find-shared-meeting-times')
    ).toBe(false);
  });

  it.each(contracts.filter(([, , , , readOnly]) => !readOnly))(
    '%s requires confirmation, encodes the owner/occurrence and preserves the body',
    async (name, method, path) => {
      const { call, graphRequest } = setup();
      const body = {
        comment: 'Approved test response',
        sendResponse: true,
        ...(name === 'accept-shared-calendar-event'
          ? {}
          : {
              proposedNewTime: {
                start: { dateTime: '2026-12-02T10:00:00', timeZone: 'Pacific Standard Time' },
                end: { dateTime: '2026-12-02T10:30:00', timeZone: 'Pacific Standard Time' },
              },
            }),
      };
      const params = { userId: owner, eventId, body };
      expect((await call(name, params)).isError).toBe(true);
      expect(graphRequest).not.toHaveBeenCalled();
      await call(name, { ...params, confirm: true });
      expect(graphRequest).toHaveBeenCalledOnce();
      expect(graphRequest.mock.calls[0][0]).toBe(`${eventPath}/${path.split('/').at(-1)}`);
      expect(graphRequest.mock.calls[0][1].method).toBe(method);
      expect(JSON.parse(graphRequest.mock.calls[0][1].body)).toEqual(body);
    }
  );

  it('preserves an explicit silent response instead of replacing false with a default', async () => {
    const { call, graphRequest } = setup();
    await call('decline-shared-calendar-event', {
      body: { sendResponse: false },
      confirm: true,
      userId: owner,
      eventId,
    });
    expect(JSON.parse(graphRequest.mock.calls[0][1].body)).toEqual({ sendResponse: false });
  });

  it('reads responseStatus with select and timezone on the owner calendar', async () => {
    const { call, graphRequest } = setup();
    graphRequest.mockResolvedValueOnce(
      jsonResult({ id: eventId, responseStatus: { response: 'declined' } })
    );
    const result = await call('get-shared-calendar-event', {
      userId: owner,
      eventId,
      select: 'id,responseStatus',
      timezone: 'Pacific Standard Time',
    });
    expect(graphRequest.mock.calls[0][0]).toBe(`${eventPath}?$select=id,responseStatus`);
    expect(graphRequest.mock.calls[0][1].headers.Prefer).toContain(
      'outlook.timezone="Pacific Standard Time"'
    );
    expect(JSON.parse(result.content[0].text).responseStatus.response).toBe('declined');
  });

  it('requires an instance time window and sends the series master to /instances', async () => {
    const { tools, call, graphRequest } = setup();
    const schema = tools.get('list-shared-calendar-event-instances')!.config
      .inputSchema as z.ZodType;
    expect(schema.safeParse({ userId: owner, eventId }).success).toBe(false);
    await call('list-shared-calendar-event-instances', {
      userId: owner,
      eventId: 'series-master',
      ...windowParams,
    });
    const url = new URL(graphRequest.mock.calls[0][0], 'https://graph.microsoft.com');
    expect(url.pathname).toBe(
      '/users/executive%2Bcalendar%40contoso.org/events/series-master/instances'
    );
    for (const [key, value] of Object.entries(windowParams))
      expect(url.searchParams.get(key)).toBe(value);
  });

  it('finds meeting times in read-only mode without a confirmation gate', async () => {
    const { call, graphRequest } = setup({ readOnly: true });
    const body = { meetingDuration: 'PT30M', maxCandidates: 3 };
    await call('find-shared-meeting-times', { userId: owner, body });
    expect(graphRequest.mock.calls[0][0]).toBe(
      '/users/executive%2Bcalendar%40contoso.org/findMeetingTimes'
    );
    expect(JSON.parse(graphRequest.mock.calls[0][1].body)).toEqual(body);
  });

  it('routes attachment bytes through the existing blob-output contract', async () => {
    vi.stubEnv('MS365_MCP_BLOB_OUTPUT', '1');
    const { call, graphRequest } = setup();
    await call('get-shared-calendar-event-attachment-content', {
      userId: owner,
      eventId,
      attachmentId: 'file/id=',
      fileName: 'agenda.pdf',
    });
    expect(graphRequest.mock.calls[0][0]).toBe(`${eventPath}/attachments/file%2Fid=/$value`);
    expect(graphRequest.mock.calls[0][1]).toMatchObject({
      forceBinary: true,
      blobOutput: { fileName: 'agenda.pdf' },
    });
  });

  it.each([
    ['list-calendar-view-delta', '/me/calendarView/delta()', windowParams],
    [
      'list-shared-calendar-view-delta',
      '/users/executive%2Bcalendar%40contoso.org/calendarView/delta()',
      { userId: owner, ...windowParams },
    ],
    [
      'list-mail-folder-messages-delta',
      '/me/mailFolders/inbox/messages/delta()',
      { mailFolderId: 'inbox' },
    ],
    [
      'list-shared-mailbox-folder-messages-delta',
      '/users/executive%2Bcalendar%40contoso.org/mailFolders/inbox/messages/delta()',
      { userId: owner, mailFolderId: 'inbox' },
    ],
  ])(
    '%s preserves the final deltaLink while merging paginated reads',
    async (name, path, params) => {
      const { call, graphRequest } = setup({ enabledTools: `^${name}$`, allowedScopes: undefined });
      const nextLink = `https://graph.microsoft.com/v1.0${path}?$skiptoken=page2`;
      const deltaLink = `https://graph.microsoft.com/v1.0${path}?$deltatoken=next-round`;
      graphRequest
        .mockResolvedValueOnce(
          jsonResult({ value: [{ id: 'first' }], '@odata.nextLink': nextLink })
        )
        .mockResolvedValueOnce(
          jsonResult({ value: [{ id: 'second' }], '@odata.deltaLink': deltaLink })
        );
      const result = await call(name, { ...params, fetchAllPages: true });
      expect(graphRequest).toHaveBeenCalledTimes(2);
      expect(graphRequest.mock.calls[1][0]).toBe(`${path}?$skiptoken=page2`);
      expect(JSON.parse(result.content[0].text)).toMatchObject({
        value: [{ id: 'first' }, { id: 'second' }],
        '@odata.deltaLink': deltaLink,
      });
    }
  );

  it('documents only the requested Tier 1 and Tier 2 additions under work', () => {
    const presets = readFileSync(new URL('../PRESETS.md', import.meta.url), 'utf8');
    for (const [name] of contracts) expect(presets).toContain(`\`${name}\``);
    const registry = buildToolsRegistry(false, true);
    expect(registry.has('forward-shared-calendar-event')).toBe(false);
    expect(registry.has('cancel-shared-calendar-event')).toBe(false);
  });

  describe.each([
    [
      'list-shared-calendar-view-delta',
      '/users/executive%2Bcalendar%40contoso.org/calendarView/delta()',
      { userId: owner },
      windowParams,
    ],
    [
      'list-shared-mailbox-folder-messages-delta',
      '/users/executive%2Bcalendar%40contoso.org/mailFolders/inbox/messages/delta()',
      { userId: owner, mailFolderId: 'inbox' },
      {},
    ],
  ])('%s incremental synchronization', (name, path, identity, initialQuery) => {
    it('exposes the optional token input in normal and discovery modes', () => {
      const { tools } = setup();
      expect(tools.get(name)?.config.inputSchema.shape.token.isOptional()).toBe(true);
      const entry = buildToolsRegistry(false, true).get(name)!;
      expect(describeToolSchema(entry.tool, entry.config).parameters).toContainEqual(
        expect.objectContaining({ name: 'token', required: false, schema: { type: 'string' } })
      );
    });

    it('resumes a second sync from the returned deltaLink without repeating initial query params', async () => {
      const { call, graphRequest } = setup();
      const cursor = 'opaque+cursor/with=padding%';
      const deltaLink = `https://graph.microsoft.com/v1.0${path}?$deltatoken=${encodeURIComponent(cursor)}`;
      graphRequest.mockResolvedValueOnce(
        jsonResult({ value: [{ id: 'initial' }], '@odata.deltaLink': deltaLink })
      );
      const first = await call(name, { ...identity, ...initialQuery });
      const firstUrl = new URL(graphRequest.mock.calls[0][0], 'https://graph.microsoft.com');
      for (const [key, value] of Object.entries(initialQuery))
        expect(firstUrl.searchParams.get(key)).toBe(value);
      const savedLink = JSON.parse(first.content[0].text)['@odata.deltaLink'];
      graphRequest.mockResolvedValueOnce(
        jsonResult({ value: [{ id: 'changed' }], '@odata.deltaLink': deltaLink })
      );
      const second = await call(name, { ...identity, token: savedLink });
      const resumed = new URL(graphRequest.mock.calls[1][0], 'https://graph.microsoft.com');
      expect(resumed.pathname).toBe(path);
      expect([...resumed.searchParams]).toEqual([['$deltatoken', cursor]]);
      expect(JSON.parse(second.content[0].text).value).toEqual([{ id: 'changed' }]);
    });

    it('continues a nextLink with skiptoken and no initial query parameters', async () => {
      const { call, graphRequest } = setup();
      const cursor = 'next+page/value=';
      await call(name, {
        ...identity,
        skiptoken: `https://graph.microsoft.com/v1.0${path}?$skiptoken=${encodeURIComponent(cursor)}`,
      });
      const url = new URL(graphRequest.mock.calls[0][0], 'https://graph.microsoft.com');
      expect(url.pathname).toBe(path);
      expect([...url.searchParams]).toEqual([['$skiptoken', cursor]]);
    });

    it('refuses a link without a delta cursor before sending a request', async () => {
      const { call, graphRequest } = setup();
      const result = await call(name, {
        ...identity,
        token: `https://graph.microsoft.com/v1.0${path}?$select=id`,
      });
      expect(result.isError).toBe(true);
      expect(JSON.parse(result.content[0].text).error).toBe('invalid_delta_token');
      expect(graphRequest).not.toHaveBeenCalled();
    });

    it('extracts only the cursor and cannot change the requested owner or request host', async () => {
      const { call, graphRequest } = setup();
      await call(name, {
        ...identity,
        token: 'https://example.invalid/users/another-owner/events?$deltatoken=opaque',
      });
      expect(graphRequest.mock.calls[0][0]).toBe(`${path}?$deltatoken=opaque`);
    });

    it('returns an expired-cursor error without silently restarting a full synchronization', async () => {
      const { call, graphRequest } = setup();
      const expired = {
        ...jsonResult({ error: { code: 'SyncStateNotFound', status: 410 } }),
        isError: true,
      };
      graphRequest.mockResolvedValueOnce(expired);
      const result = await call(name, {
        ...identity,
        token: `https://graph.microsoft.com/v1.0${path}?$deltatoken=expired`,
      });
      expect(result).toEqual(expired);
      expect(graphRequest).toHaveBeenCalledOnce();
    });
  });
});
