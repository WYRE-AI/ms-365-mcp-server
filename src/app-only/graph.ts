/* eslint-disable no-control-regex -- Deliberately reject control characters at the trust boundary. */
import { AppOnlyError, numeric } from './errors.js';
import type { Credentials } from './credentials.js';
import type { TokenProvider } from './token-provider.js';
import type { Route } from './tools.js';
import { protectContent } from './untrusted.js';
export const ID_ALLOWLIST_RE = /^[A-Za-z0-9_\-=+]{1,512}$/;

export function queryArguments(args: Record<string, unknown>): {
  query: URLSearchParams;
  headers: Record<string, string>;
} {
  const query = new URLSearchParams(),
    headers: Record<string, string> = {};
  const invalid = (detail: string): never => {
    throw new AppOnlyError(400, 'invalid_arguments', -32602, { detail });
  };
  for (const name of ['top', 'skip', 'filter', 'select', 'orderby', 'search', 'count', 'expand']) {
    if (args[name] !== undefined && args[`$${name}`] !== undefined)
      invalid(`Argument "${name}" duplicates "$${name}".`);
    const value = args[name] ?? args[`$${name}`];
    if (value === undefined) continue;
    const text = Array.isArray(value) ? value.join(',') : String(value);
    if (name === 'expand') {
      if (value === '' || (Array.isArray(value) && value.length === 0)) continue;
      invalid('Argument "expand" must be absent, "" or [].');
    }
    query.set(`$${name}`, text);
    if (name === 'count' && value === true) headers.ConsistencyLevel = 'eventual';
  }
  for (const name of ['startDateTime', 'endDateTime'])
    if (args[name] !== undefined) query.set(name, String(args[name]));
  if (args.timezone !== undefined) {
    if (/["\u0000-\u001f\u007f]/.test(String(args.timezone)))
      invalid('Argument "timezone" contains unsafe characters.');
    headers.Prefer = `outlook.timezone="${args.timezone}"`;
  }
  return { query, headers };
}
export function buildGraphUrl(
  route: Route,
  args: Record<string, unknown>,
  mailbox: string
): string {
  for (const argName of Object.values(route.pathParams)) {
    const value = args[argName];
    if (typeof value !== 'string' || !ID_ALLOWLIST_RE.test(value))
      throw new AppOnlyError(400, 'invalid_arguments', -32602, {
        detail: `Path parameter "${argName}" is outside the id allowlist.`,
      });
  }
  const { query } = queryArguments(args);
  const prefix = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/`;
  const path = route.path.replace(/\{([^}]+)\}/g, (_match, p: string) =>
    encodeURIComponent(args[route.pathParams[p]] as string)
  );
  return prefix + path.slice(1) + (query.size ? '?' + query.toString() : '');
}

export async function callGraph(
  route: Route,
  args: Record<string, unknown>,
  c: Credentials,
  token: string,
  options: {
    fetchImpl: typeof fetch;
    tokenProvider: TokenProvider;
    env: Record<string, string | undefined>;
    now: () => number;
    sleep?: (ms: number) => Promise<void>;
  }
): Promise<unknown> {
  let url = buildGraphUrl(route, args, c.mailbox);
  const { headers } = queryArguments(args);
  if (route.method === 'GET')
    headers.Prefer = [headers.Prefer, 'outlook.body-content-type="text"']
      .filter(Boolean)
      .join(', ');
  const cap = numeric(options.env, 'M365_APP_ONLY_MAX_RESPONSE_BYTES', 4000000);
  const maxPages = numeric(options.env, 'M365_APP_ONLY_MAX_PAGES', 10);
  let total = 0,
    combined: unknown[] = [],
    result: Record<string, unknown> = {};
  for (let page = 0; page < maxPages; page++) {
    let data: Record<string, unknown> = {};
    for (let attempt = 0; attempt < 2; attempt++) {
      const controller = new AbortController(),
        timer = setTimeout(() => controller.abort(), 30000);
      try {
        const response = await options.fetchImpl(url, {
          method: route.method,
          redirect: 'manual',
          signal: controller.signal,
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: 'application/json',
            ...(route.bodyArg ? { 'Content-Type': 'application/json' } : {}),
            ...headers,
          },
          ...(route.bodyArg ? { body: JSON.stringify(args[route.bodyArg]) } : {}),
        });
        if (response.status >= 300 && response.status < 400) {
          await response.body?.cancel();
          throw new Error('graph_error:redirect_refused');
        }
        if (response.status === 429 || response.status === 503) {
          const raw = response.headers.get('retry-after');
          const seconds =
            raw === null
              ? 0
              : /^\d+(?:\.\d+)?$/.test(raw)
                ? Number(raw)
                : Math.max(0, Math.ceil((Date.parse(raw) - options.now()) / 1000));
          const delay = Number.isFinite(seconds) ? seconds : 0;
          await response.body?.cancel();
          if (raw !== null && delay <= 10 && attempt === 0) {
            await (options.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms))))(
              delay * 1000
            );
            continue;
          }
          throw new Error(`graph_throttled:retry_after=${delay}`);
        }
        if (response.status === 401) {
          options.tokenProvider.drop(c);
          await response.body?.cancel();
          throw new Error('auth_failed:graph_unauthorized');
        }
        if (response.status === 403 || response.status === 404) {
          await response.body?.cancel();
          throw new Error(
            response.status === 403
              ? 'graph_forbidden:check_exchange_rbac_scope'
              : 'graph_not_found'
          );
        }
        const chunks: Uint8Array[] = [];
        const reader = response.body?.getReader();
        if (reader) {
          while (true) {
            const chunk = await reader.read();
            if (chunk.done) break;
            total += chunk.value.byteLength;
            if (total > cap) {
              controller.abort();
              await reader.cancel();
              throw new Error('graph_response_too_large');
            }
            chunks.push(chunk.value);
          }
        }
        let parsed: unknown;
        try {
          parsed = chunks.length ? JSON.parse(Buffer.concat(chunks).toString()) : { ok: true };
        } catch {
          throw new Error(`graph_error:${response.status}:invalid_json`);
        }
        if (!response.ok) {
          const error = parsed as { error?: { code?: unknown } };
          const code = String(error?.error?.code ?? '')
            .replace(/[^A-Za-z0-9_.-]/g, '')
            .slice(0, 64);
          throw new Error(`graph_error:${response.status}:${code}`);
        }
        data = parsed as Record<string, unknown>;
        break;
      } finally {
        clearTimeout(timer);
      }
    }
    result = data;
    if (Array.isArray(data.value)) combined = combined.concat(data.value);
    const next = data['@odata.nextLink'];
    delete data['@odata.nextLink'];
    if (next && args.fetchAllPages) {
      let nextUrl: URL;
      try {
        if (typeof next !== 'string') throw new Error('mailbox_mismatch');
        nextUrl = new URL(next);
        const pathname = nextUrl.pathname;
        const decoded = decodeURIComponent(pathname);
        if (
          nextUrl.origin !== 'https://graph.microsoft.com' ||
          nextUrl.username ||
          nextUrl.password ||
          /%2f|%5c/i.test(pathname) ||
          decoded.includes('..') ||
          !decoded.toLowerCase().startsWith(`/v1.0/users/${c.mailbox}/`)
        )
          throw new Error('mailbox_mismatch');
      } catch {
        throw new Error('mailbox_mismatch');
      }
      if (page + 1 < maxPages) {
        url = nextUrl.href;
        continue;
      }
    }
    if (next) result.hasMore = true;
    break;
  }
  if (Array.isArray(result.value)) result.value = combined;
  return protectContent(
    result,
    numeric(options.env, 'M365_APP_ONLY_MAX_BODY_CHARS', 20000),
    route.method === 'GET'
  );
}
