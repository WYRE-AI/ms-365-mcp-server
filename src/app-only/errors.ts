export class AppOnlyError extends Error {
  constructor(
    public status: number,
    message: string,
    public code = -32001,
    public data: Record<string, unknown> = {}
  ) {
    super(message);
  }
}
export const numeric = (
  env: Record<string, string | undefined>,
  key: string,
  fallback: number
): number => {
  const value = env[key];
  if (value === undefined) return fallback;
  const n = Number(value);
  if (!Number.isSafeInteger(n) || n <= 0) throw new Error(`Invalid ${key}`);
  return n;
};
