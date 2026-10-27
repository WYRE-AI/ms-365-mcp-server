import { createServer, type ServerResponse } from 'node:http';
import { randomUUID } from 'node:crypto';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { zodToJsonSchema } from 'zod-to-json-schema';
import { credentials, normalizeMailbox, validMailbox, type Credentials } from './credentials.js';
import { AppOnlyError, numeric } from './errors.js';
import { guardAppOnly } from './guard.js';
import { verifyS2sHeader, S2S_HEADER } from './s2s-verify.js';
import { ROUTES, type Tier } from './tools.js';
import { hardDisabled } from './hard-disabled.js';
import { SHARED_MAILBOX_TOOL_SCHEMAS, validateSharedMailboxToolArguments } from './schemas.js';
import { createTokenProvider, type TokenProvider } from './token-provider.js';
import { callGraph, queryArguments } from './graph.js';
import { createLog } from './log.js';
export interface AppOnlyOptions {
  fetchImpl?: typeof fetch;
  tokenProvider?: TokenProvider;
  env?: Record<string, string | undefined>;
  now?: () => number;
  sleep?: (ms: number) => Promise<void>;
  logOutput?: (line: string) => void;
}
export function createAppOnlyServer(options: AppOnlyOptions = {}) {
  const env = options.env ?? process.env,
    now = options.now ?? Date.now;
  const config = guardAppOnly(env, []),
    provider = options.tokenProvider ?? createTokenProvider({ now });
  const log = createLog(env, options.logOutput);
  if (!config.secret)
    log({ level: 'warn', msg: 'Insecure development mode: S2S disabled; forced 127.0.0.1' });
  const levels: Tier[] = ['read', 'write', 'admin'];
  const tier = ['true', '1'].includes(env.READ_ONLY_MODE ?? '')
    ? 'read'
    : (env.M365_APP_ONLY_MAX_TIER ?? 'write');
  if (!levels.includes(tier as Tier)) throw new Error('Invalid M365_APP_ONLY_MAX_TIER');
  const pinned =
    env.M365_APP_ONLY_PINNED_MAILBOX === undefined
      ? undefined
      : normalizeMailbox(env.M365_APP_ONLY_PINNED_MAILBOX);
  if (pinned !== undefined && !validMailbox(pinned))
    throw new Error('Invalid M365_APP_ONLY_PINNED_MAILBOX');
  const bindings = new Map<string, { mailbox: string; at: number }>();
  const bindingTtl = numeric(env, 'M365_APP_ONLY_BIND_TTL_SECONDS', 3600) * 1000;
  const allowed = (name: string): boolean =>
    Object.prototype.hasOwnProperty.call(ROUTES, name) &&
    !hardDisabled(name) &&
    levels.indexOf(ROUTES[name].tier) <= levels.indexOf(tier as Tier);
  const validate = (name: string, args: Record<string, unknown> | undefined): void => {
    if (!allowed(name))
      throw new AppOnlyError(400, 'tool_disabled_app_only', -32601, { tool: name });
    const detail = validateSharedMailboxToolArguments(name, args);
    if (detail) throw new AppOnlyError(400, 'invalid_arguments', -32602, { detail });
    queryArguments(args ?? {});
  };
  const json = (res: ServerResponse, status: number, body: unknown): void => {
    res.writeHead(status, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(body));
  };
  const server = createServer({ maxHeaderSize: 32768 }, async (req, res) => {
    const started = now();
    let id: unknown = null,
      c: Credentials | undefined,
      tool: string | undefined;
    const requestId = randomUUID();
    try {
      if (req.method === 'GET' && ['/health', '/healthz'].includes(req.url ?? '')) {
        json(res, 200, { status: 'ok' });
        return;
      }
      const proof = req.headers[S2S_HEADER];
      if (config.secret && (Array.isArray(proof) || !verifyS2sHeader(proof, config.secret))) {
        json(res, 401, { error: 'unauthorized' });
        return;
      }
      if (req.url !== '/mcp') {
        json(res, 404, { error: 'not_found' });
        return;
      }
      if (req.method !== 'POST') throw new AppOnlyError(405, 'method_not_allowed', -32600);
      if (!/^application\/json(?:\s*;|$)/i.test(req.headers['content-type'] ?? ''))
        throw new AppOnlyError(415, 'content_type_must_be_application_json', -32600);
      const cap = numeric(env, 'M365_APP_ONLY_MAX_REQUEST_BYTES', 5000000);
      let bytes = 0;
      const chunks: Buffer[] = [];
      for await (const chunk of req) {
        bytes += chunk.length;
        if (bytes > cap) throw new AppOnlyError(413, 'request_too_large', -32600);
        chunks.push(Buffer.from(chunk));
      }
      let body;
      try {
        body = JSON.parse(Buffer.concat(chunks).toString());
      } catch {
        throw new AppOnlyError(400, 'parse_error', -32700);
      }
      if (!body || typeof body !== 'object' || Array.isArray(body))
        throw new AppOnlyError(400, 'invalid_request', -32600);
      id = typeof body.id === 'string' || typeof body.id === 'number' ? body.id : null;
      c = credentials(req);
      const key = provider.key(c);
      if (pinned && pinned !== c.mailbox) throw new AppOnlyError(403, 'mailbox_mismatch');
      if (env.M365_APP_ONLY_BIND_MAILBOX !== '0') {
        for (const [k, v] of bindings) if (now() - v.at >= bindingTtl) bindings.delete(k);
        const bound = bindings.get(key);
        if (bound && bound.mailbox !== c.mailbox) throw new AppOnlyError(403, 'mailbox_mismatch');
        bindings.set(key, { mailbox: c.mailbox, at: now() });
      }
      let token: string | undefined;
      if (body.method === 'tools/call') {
        tool = typeof body.params?.name === 'string' ? body.params.name : '';
        validate(tool!, body.params?.arguments);
        token = await provider.acquire(c);
      }
      const mcp = new Server(
        { name: 'm365-app-only-shared-mailbox', version: '1.0.0' },
        { capabilities: { tools: {} } }
      );
      mcp.setRequestHandler(ListToolsRequestSchema, async () => ({
        tools: Object.entries(ROUTES)
          .filter(([name]) => allowed(name))
          .map(([name, route]) => ({
            name,
            description: route.description,
            inputSchema: zodToJsonSchema(SHARED_MAILBOX_TOOL_SCHEMAS[name], {
              $refStrategy: 'none',
            }) as { type: 'object' },
          })),
      }));
      const requestCredentials = c;
      mcp.setRequestHandler(CallToolRequestSchema, async (request) => {
        try {
          validate(request.params.name, request.params.arguments);
        } catch (e) {
          const err = e as AppOnlyError;
          throw new McpError(err.code, err.message, err.data);
        }
        const accessToken = token ?? (await provider.acquire(requestCredentials));
        try {
          const result = await callGraph(
            ROUTES[request.params.name],
            request.params.arguments ?? {},
            requestCredentials,
            accessToken,
            {
              fetchImpl: options.fetchImpl ?? fetch,
              tokenProvider: provider,
              env,
              now,
              sleep: options.sleep,
            }
          );
          return { content: [{ type: 'text', text: JSON.stringify(result) }] };
        } catch (e) {
          const message = e instanceof Error ? e.message : '';
          const safe =
            /^(mailbox_mismatch|graph_response_too_large|auth_failed:graph_unauthorized|graph_forbidden:check_exchange_rbac_scope|graph_not_found|graph_throttled:retry_after=\d+(?:\.\d+)?|graph_error:[A-Za-z0-9_.:-]+)$/.test(
              message
            )
              ? message
              : 'graph_error:request_failed';
          return {
            isError: true,
            content: [{ type: 'text', text: JSON.stringify({ error: safe }) }],
          };
        }
      });
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
        enableJsonResponse: true,
      });
      res.on('close', () => {
        void transport.close();
        void mcp.close();
      });
      await mcp.connect(transport);
      await transport.handleRequest(req, res, body);
    } catch (e) {
      const err = e instanceof AppOnlyError ? e : new AppOnlyError(500, 'internal_error', -32603);
      if (!res.headersSent)
        json(res, err.status, {
          jsonrpc: '2.0',
          id,
          error: { code: err.code, message: err.message, data: err.data },
        });
    } finally {
      log({
        msg: 'request',
        requestId,
        tool: tool && Object.prototype.hasOwnProperty.call(ROUTES, tool) ? tool : undefined,
        tier:
          tool && Object.prototype.hasOwnProperty.call(ROUTES, tool)
            ? ROUTES[tool].tier
            : undefined,
        status: res.statusCode,
        durationMs: now() - started,
        mailbox: c?.mailbox,
      });
    }
  });
  return { server, host: config.host };
}
