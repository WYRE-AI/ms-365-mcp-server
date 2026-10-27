import { createHash, createHmac, randomBytes } from 'node:crypto';
import { ConfidentialClientApplication, type Configuration } from '@azure/msal-node';
import type { Credentials } from './credentials.js';
import { AppOnlyError } from './errors.js';
export interface TokenProvider {
  key(c: Credentials): string;
  acquire(c: Credentials): Promise<string>;
  drop(c: Credentials): void;
}
export const CONSENT_FIX =
  'Remove the tenant-wide Microsoft Graph application permission(s) and admin consent from the app registration; grant mailbox access only through Exchange Online RBAC for Applications scoped to this mailbox (see README: App-only shared-mailbox mode).';
export function checkConsent(token: string, c: Credentials): void {
  let p: Record<string, unknown>;
  try {
    p = JSON.parse(Buffer.from(token.split('.')[1], 'base64url').toString());
    if (!p || typeof p !== 'object' || Array.isArray(p)) throw new Error();
  } catch {
    throw new AppOnlyError(401, 'auth_failed:unknown');
  }
  if (p.tid !== c.tenantId || (p.appid ?? p.azp) !== c.clientId)
    throw new AppOnlyError(401, 'auth_failed:tenant_or_app');
  const roles = Array.isArray(p.roles)
    ? p.roles.filter(
        (r): r is string =>
          typeof r === 'string' && /^(Mail\.|Calendars\.|MailboxSettings\.)/.test(r)
      )
    : [];
  if (roles.length)
    throw new AppOnlyError(403, 'tenant_wide_consent_detected', -32001, {
      roles,
      fix: CONSENT_FIX,
    });
}
export function mapMsalError(error: unknown): AppOnlyError {
  const e = error as { errorCode?: string; errorMessage?: string };
  const code = `${e?.errorCode ?? ''} ${e?.errorMessage ?? ''}`.match(
    /(?:AADSTS|\b)(\d{5,8})\b/
  )?.[1];
  const reason =
    code === '7000222'
      ? 'secret_expired'
      : ['7000215', '700027', '700024'].includes(code ?? '')
        ? 'invalid_client'
        : ['90002', '700016'].includes(code ?? '')
          ? 'tenant_or_app'
          : 'unknown';
  return new AppOnlyError(401, `auth_failed:${reason}`, -32001, code ? { aadsts: code } : {});
}
export function createTokenProvider(
  options: {
    now?: () => number;
    appFactory?: (
      config: Configuration
    ) => Pick<ConfidentialClientApplication, 'acquireTokenByClientCredential'>;
  } = {}
): TokenProvider {
  const salt = randomBytes(32),
    now = options.now ?? Date.now;
  const apps = new Map<
    string,
    {
      app: Pick<ConfidentialClientApplication, 'acquireTokenByClientCredential'>;
      at: number;
      verdicts: Set<string>;
    }
  >();
  const key = (c: Credentials): string =>
    createHmac('sha256', salt)
      .update(
        [
          c.tenantId,
          c.clientId,
          c.secret !== undefined ? 'secret' : 'certificate',
          c.secret ??
            `${c.thumbprint}|${createHash('sha256')
              .update(c.privateKey ?? '')
              .digest('hex')}`,
        ].join('|')
      )
      .digest('hex');
  return {
    key,
    drop(c) {
      apps.delete(key(c));
    },
    async acquire(c) {
      for (const [k, v] of apps) if (now() - v.at >= 3600000) apps.delete(k);
      const k = key(c);
      let entry = apps.get(k);
      if (!entry) {
        if (apps.size >= 500) apps.delete(apps.keys().next().value!);
        const config: Configuration = {
          auth: {
            clientId: c.clientId,
            authority: `https://login.microsoftonline.com/${c.tenantId}`,
            ...(c.secret !== undefined
              ? { clientSecret: c.secret }
              : {
                  clientCertificate: { thumbprintSha256: c.thumbprint!, privateKey: c.privateKey! },
                }),
          },
          system: { loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } },
        };
        entry = {
          app: options.appFactory
            ? options.appFactory(config)
            : new ConfidentialClientApplication(config),
          at: now(),
          verdicts: new Set(),
        };
        apps.set(k, entry);
      }
      entry.at = now();
      let token: string;
      try {
        const result = await entry.app.acquireTokenByClientCredential({
          scopes: ['https://graph.microsoft.com/.default'],
        });
        if (!result?.accessToken) throw new Error();
        token = result.accessToken;
      } catch (e) {
        throw mapMsalError(e);
      }
      const digest = createHash('sha256').update(token).digest('hex');
      if (!entry.verdicts.has(digest)) {
        checkConsent(token, c);
        entry.verdicts.clear();
        entry.verdicts.add(digest);
      }
      return token;
    },
  };
}
