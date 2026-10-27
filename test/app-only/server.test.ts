import { buildGraphUrl } from '../../src/app-only/graph.js';
import { ROUTES } from '../../src/app-only/tools.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createHmac, generateKeyPairSync } from 'node:crypto';
import { request } from 'node:http';
import fixture from './fixtures/m365-shared-mailbox-sidecar-contract.json';
import { createAppOnlyServer } from '../../src/app-only/server.js';
import { guardAppOnly } from '../../src/app-only/guard.js';
import {
  createTokenProvider,
  checkConsent,
  mapMsalError,
} from '../../src/app-only/token-provider.js';
import { protectContent } from '../../src/app-only/untrusted.js';
import { verifyS2sHeader } from '../../src/app-only/s2s-verify.js';
const tenant = '11111111-1111-1111-1111-111111111111',
  client = '22222222-2222-2222-2222-222222222222';
const secret = 'a-real-test-s2s-key';
const env = {
  MS365_MCP_APP_ONLY_SHARED_MAILBOX: '1',
  CONDUIT_S2S_SECRET: secret,
  M365_APP_ONLY_MAX_TIER: 'admin',
};
const proof = (t = Math.floor(Date.now() / 1000)) =>
  `t=${t},v1=${createHmac('sha256', secret).update(`t=${t}`).digest('hex')}`;
const headers = () => ({
  'content-type': 'application/json',
  accept: 'application/json, text/event-stream',
  'x-gateway-s2s': proof(),
  'x-m365-tenant-id': tenant,
  'x-m365-client-id': client,
  'x-m365-client-secret': 'credential-secret',
  'x-m365-mailbox': 'shared@acme.com',
});
const servers: ReturnType<typeof createAppOnlyServer>['server'][] = [];
afterEach(async () => {
  for (const server of servers.splice(0))
    await new Promise<void>((resolve) => server.close(() => resolve()));
});
async function setup(
  overrides = {},
  response?: (url: string, init: Parameters<typeof fetch>[1]) => Response
) {
  const fetchImpl = vi.fn(async (url, init) =>
    response ? response(String(url), init) : new Response(JSON.stringify({ value: [] }))
  );
  const tokenProvider = {
    key: () => 'credential-key',
    acquire: vi.fn(async () => 'private-access-token'),
    drop: vi.fn(),
  };
  const logs: string[] = [];
  const app = createAppOnlyServer({
    env: { ...env, ...overrides },
    fetchImpl,
    tokenProvider,
    sleep: async () => {},
    logOutput: (line) => logs.push(line),
  });
  servers.push(app.server);
  await new Promise<void>((resolve) => app.server.listen(0, '127.0.0.1', resolve));
  const address = app.server.address() as { port: number };
  const post = (
    body: unknown,
    patch: Record<string, string | string[] | undefined> = {},
    path = '/mcp',
    method = 'POST'
  ) =>
    new Promise<{ status: number; body: any }>((resolve, reject) => {
      const h: Record<string, string | string[]> = { ...headers() };
      for (const [k, v] of Object.entries(patch)) {
        if (v === undefined) delete h[k];
        else h[k] = v;
      }
      const req = request(
        { host: '127.0.0.1', port: address.port, path, method, headers: h },
        (res) => {
          const chunks: Buffer[] = [];
          res.on('data', (c) => chunks.push(c));
          res.on('end', () =>
            resolve({ status: res.statusCode!, body: JSON.parse(Buffer.concat(chunks).toString()) })
          );
        }
      );
      req.on('error', reject);
      req.end(method === 'POST' ? JSON.stringify(body) : undefined);
    });
  const call = (name: string, args: unknown = {}, patch = {}) =>
    post({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }, patch);
  return { post, call, fetchImpl, tokenProvider, logs };
}
const rpc = (method: string) => ({ jsonrpc: '2.0', id: 1, method });
describe('contract', () => {
  it('rejects every fixture bypass before token acquisition', async () => {
    const s = await setup();
    for (const entry of fixture.mustReject) {
      const r = await s.call(entry.tool, entry.args);
      expect(r.body.error.message, JSON.stringify(entry)).toBe('invalid_arguments');
      expect(r.body.error.code).toBe(-32602);
    }
    expect(s.fetchImpl).not.toHaveBeenCalled();
    expect(s.tokenProvider.acquire).not.toHaveBeenCalled();
  });
  it('routes all 37 valid examples exactly once and encodes IDs', async () => {
    const s = await setup();
    for (const tool of fixture.tools) {
      const r = await s.call(tool.name, tool.validExample);
      expect(r.body.error, tool.name).toBeUndefined();
      expect(r.body.result.isError, tool.name).not.toBe(true);
    }
    expect(s.fetchImpl).toHaveBeenCalledTimes(37);
    for (const [url] of s.fetchImpl.mock.calls)
      expect(url).toMatch(/^https:\/\/graph.microsoft.com\/v1.0\/users\/shared%40acme.com\//);
    await s.call('get-mail-message', { messageId: 'a=+_' });
    expect(s.fetchImpl.mock.lastCall?.[0]).toContain('/messages/a%3D%2B_');
  });
  it('advertises exactly the fixture with strict schemas and no resources/prompts', async () => {
    const s = await setup();
    const r = await s.post(rpc('tools/list'));
    expect(r.body.result.tools.map((t: any) => t.name).sort()).toEqual(
      fixture.tools.map((t) => t.name).sort()
    );
    for (const t of r.body.result.tools) expect(t.inputSchema.additionalProperties).toBe(false);
    for (const name of fixture.hardDisabledTools) {
      const rejected = await s.call(name);
      expect(rejected.body.error).toMatchObject({
        code: -32601,
        message: 'tool_disabled_app_only',
      });
    }
    for (const method of ['resources/list', 'prompts/list'])
      expect((await s.post(rpc(method))).body.error.code).toBe(-32601);
    const init = await s.post({
      ...rpc('initialize'),
      params: {
        protocolVersion: '2025-03-26',
        capabilities: {},
        clientInfo: { name: 'test', version: '1' },
      },
    });
    expect(init.body.result.capabilities).toEqual({ tools: {} });
    expect(s.tokenProvider.acquire).not.toHaveBeenCalled();
  });
  it('gates default write and read-only tiers', async () => {
    for (const overrides of [{ M365_APP_ONLY_MAX_TIER: undefined }, { READ_ONLY_MODE: '1' }]) {
      const s = await setup(overrides);
      const names = (await s.post(rpc('tools/list'))).body.result.tools.map((t: any) => t.name);
      expect(names).toEqual(
        fixture.tools
          .filter((t) => (overrides.READ_ONLY_MODE ? t.tier === 'read' : t.tier !== 'admin'))
          .map((t) => t.name)
      );
      expect((await s.call('send-mail', {})).body.error.message).toBe('tool_disabled_app_only');
    }
  });
});
describe('contract round 3-5', () => {
  it('routes every mustAccept filter as exactly one query value', async () => {
    const s = await setup();
    for (const entry of (fixture as any).mustAccept) {
      const r = await s.call(entry.tool, entry.args);
      expect(r.body.error, JSON.stringify(entry)).toBeUndefined();
      expect(r.body.result.isError).not.toBe(true);
      const url = new URL(s.fetchImpl.mock.lastCall![0]);
      expect(url.searchParams.getAll('$filter')).toEqual([entry.args.filter]);
      expect(url.searchParams.has('$expand')).toBe(false);
    }
    expect(s.fetchImpl).toHaveBeenCalledTimes((fixture as any).mustAccept.length);
  });

  it.each([
    "subject eq 'a&b'",
    "contains(subject,'50%')",
    "contains(subject,'#123')",
    "contains(subject,'$expand=attachments')",
  ])('encodes the filter exactly once: %s', (filter) => {
    const raw = buildGraphUrl(ROUTES['list-mail-messages'], { filter }, 'shared@acme.com');
    const url = new URL(raw);
    expect(url.searchParams.getAll('$filter')).toEqual([filter]);
    expect([...url.searchParams.keys()]).toEqual(['$filter']);
    expect(raw.split('?')[1]).not.toContain('&');
    expect(raw).not.toContain('#');
  });

  it.each(['../x', 'a%2Fb', 'a/b', '.', 'ａ', 'x y', '', undefined, 123])(
    'rejects path parameters outside the allowlist: %s',
    (messageId) => {
      expect(() =>
        buildGraphUrl(ROUTES['get-mail-message'], { messageId }, 'shared@acme.com')
      ).toThrowError(
        expect.objectContaining({
          message: 'invalid_arguments',
          code: -32602,
          status: 400,
          data: { detail: 'Path parameter "messageId" is outside the id allowlist.' },
        })
      );
    }
  );
  it.each(['AAMkAG-_=+', 'inbox'])('accepts allowlisted IDs: %s', (messageId) => {
    expect(buildGraphUrl(ROUTES['get-mail-message'], { messageId }, 'shared@acme.com')).toBe(
      `https://graph.microsoft.com/v1.0/users/shared%40acme.com/messages/${encodeURIComponent(messageId)}`
    );
  });

  it('preserves recipient addresses and rejects display forms and whitespace', async () => {
    const s = await setup();
    const example = (fixture.tools.find((t) => t.name === 'create-draft-email') as any)
      .validExample;
    const args = (address: string) => ({
      ...example,
      body: { ...example.body, toRecipients: [{ emailAddress: { address } }] },
    });
    const address = 'Ops.Team+x@Acme.com';
    const r = await s.call('create-draft-email', args(address));
    expect(r.body.error).toBeUndefined();
    expect(r.body.result.isError).not.toBe(true);
    expect(JSON.parse(s.fetchImpl.mock.lastCall![1].body).toRecipients).toEqual([
      { emailAddress: { address } },
    ]);
    s.fetchImpl.mockClear();
    for (const invalid of ['Name <a@acme.com>', ' a@acme.com'])
      expect((await s.call('create-draft-email', args(invalid))).body.error.message).toBe(
        'invalid_arguments'
      );
    expect(s.fetchImpl).not.toHaveBeenCalled();
  });

  it('omits empty expand values and rejects non-empty expand', async () => {
    const s = await setup();
    for (const expand of ['', []]) {
      const r = await s.call('list-mail-messages', { expand });
      expect(r.body.error).toBeUndefined();
      expect(r.body.result.isError).not.toBe(true);
      expect(new URL(s.fetchImpl.mock.lastCall![0]).searchParams.has('$expand')).toBe(false);
      expect(
        new URL(
          buildGraphUrl(ROUTES['list-mail-messages'], { expand }, 'shared@acme.com')
        ).searchParams.has('$expand')
      ).toBe(false);
    }
    s.fetchImpl.mockClear();
    expect((await s.call('list-mail-messages', { expand: 'attachments' })).body.error.message).toBe(
      'invalid_arguments'
    );
    expect(() =>
      buildGraphUrl(ROUTES['list-mail-messages'], { expand: 'attachments' }, 'shared@acme.com')
    ).toThrow('invalid_arguments');
    expect(s.fetchImpl).not.toHaveBeenCalled();
  });

  it('rejects a JSON string body without parsing it a second time', async () => {
    const s = await setup();
    const r = await s.call('create-draft-email', { body: JSON.stringify({ subject: 'Hi' }) });
    expect(r.body.error.message).toBe('invalid_arguments');
    expect(s.fetchImpl).not.toHaveBeenCalled();
  });
});
describe('HTTP boundaries', () => {
  it('authenticates before dispatch, health remains independent', async () => {
    const s = await setup();
    for (const v of [undefined, 'bad', proof(1)])
      expect((await s.post(rpc('tools/list'), { 'x-gateway-s2s': v })).status).toBe(401);
    expect((await s.post({}, { 'x-gateway-s2s': undefined }, '/health', 'GET')).body).toEqual({
      status: 'ok',
    });
    expect(s.fetchImpl).not.toHaveBeenCalled();
    expect(s.tokenProvider.acquire).not.toHaveBeenCalled();
    expect((await s.post({}, {}, '/else')).status).toBe(404);
    expect((await s.post({}, {}, '/mcp', 'DELETE')).status).toBe(405);
    expect((await s.post([])).status).toBe(400);
    expect((await s.post({}, { 'content-type': 'text/plain' })).status).toBe(415);
  });
  it('rejects malformed credentials and normalizes mailbox', async () => {
    const s = await setup();
    const cases = [
      ['x-m365-tenant-id', 'bad', 'tenant_id', 400],
      ['x-m365-client-id', 'bad', 'client_id', 400],
      ['x-m365-client-secret', '', 'client_secret', 400],
      ['x-m365-client-secret', 'bad\tsecret', 'client_secret', 400],
      ['x-m365-mailbox', undefined, 'mailbox', 401],
      ['x-m365-mailbox', '', 'mailbox', 401],
      ['x-m365-mailbox', 'disabled', 'mailbox', 400],
      ['x-m365-mailbox', 'x@y', 'mailbox', 400],
      ['x-m365-client-cert-key', 'YQ==', 'client_secret_and_certificate', 400],
    ] as const;
    for (const [key, value, field, status] of cases) {
      const r = await s.call('list-mail-messages', {}, { [key]: value });
      expect(r.status).toBe(status);
      expect(r.body.error.message).toBe(`bad_credentials:${field}`);
    }
    for (const field of ['tenant-id', 'client-id', 'client-secret', 'mailbox'])
      expect(
        (await s.call('list-mail-messages', {}, { [`x-m365-${field}`]: ['a', 'b'] })).status
      ).toBe(400);
    expect(
      (await s.call('list-mail-messages', {}, { 'x-m365-mailbox': ' Shared@Acme.COM ' })).body
        .result.isError
    ).not.toBe(true);
    expect(s.fetchImpl.mock.lastCall?.[0]).toContain('shared%40acme.com');
  });
  it('accepts PKCS8 certificate and rejects half or malformed pairs', async () => {
    const s = await setup();
    const { privateKey } = generateKeyPairSync('rsa', {
      modulusLength: 2048,
      privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
      publicKeyEncoding: { type: 'spki', format: 'pem' },
    });
    const h = {
      'x-m365-client-secret': undefined,
      'x-m365-client-cert-key': Buffer.from(privateKey).toString('base64'),
      'x-m365-client-cert-thumbprint': 'ab'.repeat(32),
    };
    expect((await s.call('list-mail-messages', {}, h)).status).toBe(200);
    for (const field of ['key', 'thumbprint']) {
      const r = await s.call(
        'list-mail-messages',
        {},
        { ...h, [`x-m365-client-cert-${field}`]: undefined }
      );
      expect(r.body.error.message).toBe(`bad_credentials:client_cert_${field}`);
    }
  });
  it('locks pinned and bound mailboxes before Graph', async () => {
    for (const overrides of [{ M365_APP_ONLY_PINNED_MAILBOX: 'other@acme.com' }, {}]) {
      const s = await setup(overrides);
      if (!('M365_APP_ONLY_PINNED_MAILBOX' in overrides)) await s.post(rpc('tools/list'));
      const r = await s.call('list-mail-messages', {}, { 'x-m365-mailbox': 'else@acme.com' });
      expect(r.status).toBe(403);
      expect(r.body.error.message).toBe('mailbox_mismatch');
      expect(s.fetchImpl).not.toHaveBeenCalled();
      expect(s.tokenProvider.acquire).not.toHaveBeenCalled();
    }
  });
  it('refuses aliases, attachment expansion and unsafe timezone before token calls', async () => {
    const s = await setup();
    for (const args of [{ top: 1, $top: 2 }, { expand: 'ATTACHMENTS' }, { timezone: 'UTC"' }]) {
      const r = await s.call('list-calendar-events', args);
      expect(r.body.error.message).toBe('invalid_arguments');
    }
    expect(s.tokenProvider.acquire).not.toHaveBeenCalled();
  });
  it('caps request bodies', async () => {
    const s = await setup({ M365_APP_ONLY_MAX_REQUEST_BYTES: '10' });
    expect((await s.post(rpc('tools/list'))).status).toBe(413);
  });
});
describe('Graph', () => {
  for (const [status, error] of [
    [403, 'graph_forbidden:check_exchange_rbac_scope'],
    [404, 'graph_not_found'],
    [401, 'auth_failed:graph_unauthorized'],
    [302, 'graph_error:redirect_refused'],
  ] as const)
    it(`maps ${status}`, async () => {
      const s = await setup({}, () => new Response('', { status }));
      const r = await s.call('list-mail-messages');
      expect(JSON.parse(r.body.result.content[0].text).error).toBe(error);
      if (status === 401) expect(s.tokenProvider.drop).toHaveBeenCalled();
    });
  for (const status of [429, 503])
    for (const seconds of ['0', '11'])
      it(`throttles ${status}/${seconds}`, async () => {
        const s = await setup(
          {},
          () => new Response('', { status, headers: { 'retry-after': seconds } })
        );
        expect(
          JSON.parse((await s.call('list-mail-messages')).body.result.content[0].text).error
        ).toBe(`graph_throttled:retry_after=${seconds}`);
        expect(s.fetchImpl).toHaveBeenCalledTimes(seconds === '0' ? 2 : 1);
      });
  it('caps responses and blocks cross-mailbox paging', async () => {
    const s = await setup({ M365_APP_ONLY_MAX_RESPONSE_BYTES: '5' }, () => new Response('123456'));
    expect(JSON.parse((await s.call('list-mail-messages')).body.result.content[0].text).error).toBe(
      'graph_response_too_large'
    );
    const p = await setup(
      {},
      () =>
        new Response(
          JSON.stringify({
            value: [],
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users/other%40acme.com/messages',
          })
        )
    );
    expect(
      JSON.parse(
        (await p.call('list-mail-messages', { fetchAllPages: true })).body.result.content[0].text
      ).error
    ).toBe('mailbox_mismatch');
    expect(p.fetchImpl).toHaveBeenCalledTimes(1);
  });
  it('follows an uppercase same-mailbox nextLink with an unencoded @', async () => {
    const next = 'https://graph.microsoft.com/V1.0/USERS/SHARED@ACME.COM/messages?$skip=10';
    let page = 0;
    const s = await setup(
      {},
      () =>
        new Response(
          JSON.stringify({
            value: [{ id: ++page }],
            ...(page === 1 ? { '@odata.nextLink': next } : {}),
          })
        )
    );
    const result = (await s.call('list-mail-messages', { fetchAllPages: true })).body.result;
    expect(result.isError).not.toBe(true);
    expect(JSON.parse(result.content[0].text).value).toEqual([{ id: 1 }, { id: 2 }]);
    expect(s.fetchImpl).toHaveBeenCalledTimes(2);
    expect(s.fetchImpl.mock.lastCall?.[0]).toBe(next);
  });
  it.each([
    'https://graph.microsoft.com/v1.0/users/other@acme.com/messages',
    'https://example.com/v1.0/users/shared@acme.com/messages',
    'https://graph.microsoft.com/v1.0/users/shared@acme.com/%2e%2e/other@acme.com/messages',
    'https://graph.microsoft.com/v1.0/users/shared@acme.com/messages%2fother',
    'https://graph.microsoft.com/v1.0/users/shared@acme.com/messages%5Cother',
    'https://user:password@graph.microsoft.com/v1.0/users/shared@acme.com/messages',
    'https://graph.microsoft.com/v1.0/users/shared@acme.com/messages%ZZ',
    'invalid-url',
  ])('rejects unsafe nextLink %s', async (next) => {
    const s = await setup(
      {},
      () => new Response(JSON.stringify({ value: [], '@odata.nextLink': next }))
    );
    const result = (await s.call('list-mail-messages', { fetchAllPages: true })).body.result;
    expect(result.isError).toBe(true);
    expect(JSON.parse(result.content[0].text).error).toBe('mailbox_mismatch');
    expect(s.fetchImpl).toHaveBeenCalledTimes(1);
  });
  it('hides nextLink without paging and redacts logs', async () => {
    const s = await setup(
      {},
      () =>
        new Response(
          JSON.stringify({
            body: { contentType: 'html', content: '<p>private body</p>' },
            '@odata.nextLink': 'https://private',
          })
        )
    );
    const r = await s.call('get-mail-message', { messageId: 'abc' });
    const content = JSON.parse(r.body.result.content[0].text);
    expect(content.hasMore).toBe(true);
    expect(content['@odata.nextLink']).toBeUndefined();
    expect(content.body.content).toContain('untrusted-content');
    for (const value of [
      'credential-secret',
      'private-access-token',
      'private body',
      'shared@acme.com',
    ])
      expect(s.logs.join('')).not.toContain(value);
  });
});
describe('pure security guards', () => {
  it('fails closed on startup arguments/env and placeholder secrets', () => {
    expect(() => guardAppOnly({}, [])).toThrow();
    for (const value of ['', 'secret', 'changeme', 'placeholder'])
      expect(() => guardAppOnly({ ...env, CONDUIT_S2S_SECRET: value }, [])).toThrow();
    expect(
      guardAppOnly(
        { ...env, CONDUIT_S2S_SECRET: '', MCP_ALLOW_INSECURE_DEV: '1', MCP_HTTP_HOST: '0.0.0.0' },
        []
      ).host
    ).toBe('127.0.0.1');
    for (const arg of [
      '--extra-scopes',
      '--trust-proxy-auth',
      '--enable-auth-tools',
      '--org-mode',
      '--http',
    ])
      expect(() => guardAppOnly(env, [arg])).toThrow();
    for (const key of [
      'MS365_MCP_EXTRA_SCOPES',
      'MS365_MCP_CLIENT_SECRET',
      'MS365_MCP_OBO_TEST',
      'MS365_MCP_OAUTH_TEST',
    ])
      expect(() => guardAppOnly({ ...env, [key]: 'yes' }, [])).toThrow();
    expect(verifyS2sHeader(proof(), secret)).toBe(true);
  });
  const c = {
    tenantId: tenant,
    clientId: client,
    mailbox: 'shared@acme.com',
    secret: 'cache-secret',
  };
  const jwt = (p: unknown) => `a.${Buffer.from(JSON.stringify(p)).toString('base64url')}.c`;
  it('checks tenant, client and tenant-wide consent', () => {
    expect(() => checkConsent(jwt({ tid: tenant, appid: client, roles: [] }), c)).not.toThrow();
    try {
      checkConsent(jwt({ tid: tenant, appid: client, roles: ['Mail.Read'] }), c);
      throw new Error('expected denial');
    } catch (e: any) {
      expect(e.status).toBe(403);
      expect(e.data.roles).toEqual(['Mail.Read']);
      expect(e.data.fix).toContain('Exchange Online RBAC');
    }
    expect(() => checkConsent(jwt({ tid: 'wrong', appid: client }), c)).toThrow(
      'auth_failed:tenant_or_app'
    );
  });
  it('maps AADSTS without leaking raw errors', () => {
    for (const [code, reason] of [
      ['7000222', 'secret_expired'],
      ['7000215', 'invalid_client'],
      ['90002', 'tenant_or_app'],
      ['700016', 'tenant_or_app'],
      ['700027', 'invalid_client'],
      ['700024', 'invalid_client'],
    ])
      expect(mapMsalError({ errorMessage: `AADSTS${code}: private` }).message).toBe(
        `auth_failed:${reason}`
      );
  });
  it('reuses apps, hashes credential keys, expires and evicts', async () => {
    let now = 0;
    const factory = vi.fn(() => ({
      acquireTokenByClientCredential: vi.fn(async () => ({
        accessToken: jwt({ tid: tenant, appid: client }),
      })),
    }));
    const p = createTokenProvider({ now: () => now, appFactory: factory as any });
    expect(p.key(c)).not.toContain(c.secret);
    expect(p.key({ ...c, secret: 'different' })).not.toBe(p.key(c));
    await p.acquire(c);
    await p.acquire(c);
    expect(factory).toHaveBeenCalledTimes(1);
    now = 3600001;
    await p.acquire(c);
    expect(factory).toHaveBeenCalledTimes(2);
  });
  it('strips html and marker injection, truncates and shares nonce', () => {
    const r = protectContent(
      {
        body: { contentType: 'html', content: '<b>&amp;untrusted-content abcdef</b>' },
        bodyPreview: 'preview',
      },
      4
    ) as any;
    expect(r.body.contentType).toBe('text');
    expect(r.body.content).toContain('[truncated ');
    expect(r.body.content).not.toContain('<b>');
    const nonce = r.body.content.match(/nonce="([a-f0-9]+)"/)[1];
    expect(r.bodyPreview).toContain(nonce);
    expect(r._untrusted_notice).toBeTruthy();
  });
});

describe('additional security boundaries', () => {
  it('returns consent and authentication failures as HTTP errors before Graph', async () => {
    const s = await setup();
    for (const [payload, status, message] of [
      [{ tid: tenant, appid: client, roles: ['Mail.Read'] }, 403, 'tenant_wide_consent_detected'],
      [{ tid: 'wrong', appid: client, roles: [] }, 401, 'auth_failed:tenant_or_app'],
    ] as const) {
      s.tokenProvider.acquire.mockImplementation(async () => {
        checkConsent(`a.${Buffer.from(JSON.stringify(payload)).toString('base64url')}.c`, {
          tenantId: tenant,
          clientId: client,
          mailbox: 'shared@acme.com',
        });
        return 'unused';
      });
      const r = await s.call('list-mail-messages');
      expect(r.status).toBe(status);
      expect(r.body.error.message).toBe(message);
      if (status === 403) {
        expect(r.body.error.data.roles).toEqual(['Mail.Read']);
        expect(r.body.error.data.fix).toContain('Exchange Online RBAC');
      }
    }
    expect(s.fetchImpl).not.toHaveBeenCalled();
  });
  it('enforces secret and certificate format limits without acquiring tokens', async () => {
    const s = await setup();
    for (const patch of [
      { 'x-m365-client-secret': undefined },
      { 'x-m365-client-secret': 'x'.repeat(1025) },
      {
        'x-m365-client-secret': undefined,
        'x-m365-client-cert-key': 'bad!',
        'x-m365-client-cert-thumbprint': 'ab'.repeat(32),
      },
      {
        'x-m365-client-secret': undefined,
        'x-m365-client-cert-key': Buffer.from('-----BEGIN PRIVATE KEY-----\ninvalid').toString(
          'base64'
        ),
        'x-m365-client-cert-thumbprint': 'ab'.repeat(32),
      },
      { 'x-m365-mailbox': 'none' },
      { 'x-m365-mailbox': '-' },
      { 'x-m365-mailbox': 'x@@acme.com' },
      { 'x-m365-mailbox': 'a'.repeat(250) + '@acme.com' },
    ])
      expect((await s.call('list-mail-messages', {}, patch)).status).toBe(400);
    expect(s.tokenProvider.acquire).not.toHaveBeenCalled();
  });
  it('keeps secrets, certificate key, token, arguments and body out of request logs', async () => {
    const s = await setup(
      {},
      () =>
        new Response(
          JSON.stringify({ body: { contentType: 'text', content: 'highly confidential body' } })
        )
    );
    const { privateKey } = generateKeyPairSync('rsa', {
      modulusLength: 2048,
      privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
      publicKeyEncoding: { type: 'spki', format: 'pem' },
    });
    const encoded = Buffer.from(privateKey).toString('base64');
    await s.call('list-mail-messages');
    await s.call(
      'get-mail-message',
      { messageId: 'sensitive-id' },
      {
        'x-m365-client-secret': undefined,
        'x-m365-client-cert-key': encoded,
        'x-m365-client-cert-thumbprint': 'ab'.repeat(32),
      }
    );
    const output = s.logs.join('');
    for (const value of [
      'credential-secret',
      privateKey,
      encoded,
      'private-access-token',
      'highly confidential body',
      'sensitive-id',
      'shared@acme.com',
    ])
      expect(output).not.toContain(value);
    expect(JSON.parse(s.logs[0]).mailboxHash).toMatch(/^[a-f0-9]{12}$/);
  });
  it('follows bounded same-mailbox pages and combines values', async () => {
    let page = 0;
    const s = await setup(
      { M365_APP_ONLY_MAX_PAGES: '2' },
      () =>
        new Response(
          JSON.stringify({
            value: [{ id: ++page }],
            '@odata.nextLink': `https://graph.microsoft.com/v1.0/users/shared%40acme.com/messages?$skip=${page}`,
          })
        )
    );
    const r = await s.call('list-mail-messages', { fetchAllPages: true });
    const result = JSON.parse(r.body.result.content[0].text);
    expect(result.value).toEqual([{ id: 1 }, { id: 2 }]);
    expect(result.hasMore).toBe(true);
    expect(s.fetchImpl).toHaveBeenCalledTimes(2);
  });
  it('rejects normalized traversal in a nextLink', async () => {
    const s = await setup(
      {},
      () =>
        new Response(
          JSON.stringify({
            value: [],
            '@odata.nextLink':
              'https://graph.microsoft.com/v1.0/users/shared%40acme.com/../other%40acme.com/messages',
          })
        )
    );
    expect(
      JSON.parse(
        (await s.call('list-mail-messages', { fetchAllPages: true })).body.result.content[0].text
      ).error
    ).toBe('mailbox_mismatch');
    expect(s.fetchImpl).toHaveBeenCalledTimes(1);
  });
  it('builds typed query parameters and Prefer headers', async () => {
    const s = await setup();
    await s.call('list-calendar-events', {
      select: ['id', 'subject'],
      orderby: ['subject'],
      count: true,
      search: '"phrase"',
      timezone: 'UTC',
    });
    const [url, init] = s.fetchImpl.mock.lastCall!;
    const parsed = new URL(url);
    expect(parsed.searchParams.get('$select')).toBe('id,subject');
    expect(parsed.searchParams.get('$search')).toBe('"phrase"');
    expect(init.headers.ConsistencyLevel).toBe('eventual');
    expect(init.headers.Prefer).toContain('outlook.timezone="UTC"');
    expect(init.redirect).toBe('manual');
  });
  it('retries a throttled request only once and permits a subsequent success', async () => {
    let calls = 0;
    const s = await setup({}, () =>
      ++calls === 1
        ? new Response('', {
            status: 503,
            headers: { 'retry-after': new Date(Date.now() - 1000).toUTCString() },
          })
        : new Response(null, { status: 204 })
    );
    const result = (await s.call('delete-mail-message', { messageId: 'abc' })).body.result;
    expect(result.isError).not.toBe(true);
    expect(JSON.parse(result.content[0].text)).toEqual({ ok: true });
    expect(s.fetchImpl).toHaveBeenCalledTimes(2);
  });
  it('bounds the app cache and supports explicit invalidation', async () => {
    const c = { tenantId: tenant, clientId: client, mailbox: 'shared@acme.com', secret: 'first' };
    const token = `a.${Buffer.from(JSON.stringify({ tid: tenant, appid: client })).toString('base64url')}.c`;
    const factory = vi.fn(() => ({
      acquireTokenByClientCredential: vi.fn(async () => ({ accessToken: token })),
    }));
    const provider = createTokenProvider({ appFactory: factory as any });
    await provider.acquire(c);
    for (let i = 0; i < 500; i++) await provider.acquire({ ...c, secret: `secret-${i}` });
    await provider.acquire(c);
    expect(factory).toHaveBeenCalledTimes(502);
    provider.drop(c);
    await provider.acquire(c);
    expect(factory).toHaveBeenCalledTimes(503);
  });
});
