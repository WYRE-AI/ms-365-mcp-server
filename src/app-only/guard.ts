export const FORBIDDEN_ENV = [
  'MS365_MCP_EXTRA_SCOPES',
  'MS365_MCP_TRUST_PROXY_AUTH',
  'MS365_MCP_OBO',
  'MS365_MCP_OBO_ALLOWED_TENANTS',
  'MS365_MCP_OBO_ALLOWED_USERS',
  'MS365_MCP_CLIENT_SECRET',
  'MS365_MCP_OAUTH_CLIENT_ID',
  'MS365_MCP_OAUTH_CLIENT_SECRET',
  'MS365_MCP_OAUTH_REDIRECT_URI',
] as const;
export const enabled = (value: string | undefined): boolean => value === '1' || value === 'true';
export function guardAppOnly(
  env: Record<string, string | undefined>,
  args: string[]
): { host: string; secret: string } {
  if (!enabled(env.MS365_MCP_APP_ONLY_SHARED_MAILBOX))
    throw new Error('MS365_MCP_APP_ONLY_SHARED_MAILBOX must be 1 or true');
  if (args.some((a) => a !== '--app-only-shared-mailbox'))
    throw new Error('App-only mode refuses other CLI arguments');
  for (const key of Object.keys(env)) {
    if (
      (FORBIDDEN_ENV.includes(key as (typeof FORBIDDEN_ENV)[number]) ||
        key.startsWith('MS365_MCP_OBO') ||
        key.startsWith('MS365_MCP_OAUTH_')) &&
      env[key] &&
      !['0', 'false'].includes(env[key]!.toLowerCase())
    )
      throw new Error(`App-only mode refuses ${key}`);
  }
  if (
    env.NODE_ENV?.toLowerCase() !== 'development' &&
    env.MCP_ALLOW_INSECURE_DEV &&
    !['0', 'false'].includes(env.MCP_ALLOW_INSECURE_DEV.toLowerCase())
  )
    throw new Error('MCP_ALLOW_INSECURE_DEV requires NODE_ENV=development');
  const secret = (env.CONDUIT_S2S_SECRET ?? '').trim();
  if (
    [
      'replace-with-a-real-secret',
      'your-s2s-secret',
      'your-secret',
      'changeme',
      'change-me',
      'placeholder',
      'secret',
      'example',
      'replace-me',
      's2s-secret',
    ].includes(secret.toLowerCase())
  )
    throw new Error('CONDUIT_S2S_SECRET is a known placeholder');
  if (secret && secret.length < 32)
    throw new Error('CONDUIT_S2S_SECRET must be at least 32 characters');
  if (!secret && env.MCP_ALLOW_INSECURE_DEV !== '1') throw new Error('CONDUIT_S2S_SECRET is empty');
  return { secret, host: !secret ? '127.0.0.1' : env.MCP_HTTP_HOST || '0.0.0.0' };
}
