import { createHash } from 'node:crypto';
export function createLog(
  env: Record<string, string | undefined>,
  write: (line: string) => void = (line) => process.stdout.write(line)
) {
  return (fields: {
    level?: string;
    msg: string;
    tool?: string;
    tier?: string;
    status?: number;
    durationMs?: number;
    requestId?: string;
    mailbox?: string;
  }): void => {
    const { mailbox, ...safe } = fields;
    const levels = ['debug', 'info', 'warn', 'error', 'silent'];
    const threshold = levels.indexOf(env.M365_APP_ONLY_LOG_LEVEL ?? 'info');
    if (levels.indexOf(fields.level ?? 'info') < Math.max(0, threshold)) return;
    write(
      JSON.stringify({
        ts: new Date().toISOString(),
        level: 'info',
        ...safe,
        ...(mailbox
          ? {
              mailboxHash: createHash('sha256').update(mailbox).digest('hex').slice(0, 12),
              domain: mailbox.split('@')[1],
            }
          : {}),
      }) + '\n'
    );
  };
}
