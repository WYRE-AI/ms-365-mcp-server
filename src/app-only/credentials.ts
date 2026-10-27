import { createPrivateKey } from 'node:crypto';
import type { IncomingMessage } from 'node:http';
import { AppOnlyError } from './errors.js';
export interface Credentials {
  tenantId: string;
  clientId: string;
  mailbox: string;
  secret?: string;
  privateKey?: string;
  thumbprint?: string;
}
export function normalizeMailbox(value: string): string {
  return value.trim().toLowerCase().normalize('NFC');
}
export function validMailbox(value: string): boolean {
  return (
    value.length <= 254 &&
    /^[a-z0-9._%+-]+@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z]{2,}$/.test(value)
  );
}
export function credentials(req: IncomingMessage): Credentials {
  const bad = (field: string, status = 400): never => {
    throw new AppOnlyError(status, `bad_credentials:${field}`);
  };
  const header = (field: string): string | undefined => {
    const key = `x-m365-${field.replace(/_/g, '-')}`;
    if (
      req.rawHeaders.filter((v, i) => i % 2 === 0 && v.toLowerCase() === key).length > 1 ||
      Array.isArray(req.headers[key])
    )
      bad(field);
    return req.headers[key] as string | undefined;
  };
  const guid = (field: string): string => {
    const value = header(field)?.trim().toLowerCase();
    if (!value || !/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(value)) bad(field);
    return value!;
  };
  const tenantId = guid('tenant_id'),
    clientId = guid('client_id');
  const secret = header('client_secret'),
    key = header('client_cert_key'),
    thumb = header('client_cert_thumbprint');
  if (secret !== undefined && (key !== undefined || thumb !== undefined))
    bad('client_secret_and_certificate');
  let privateKey: string | undefined, thumbprint: string | undefined;
  if (key !== undefined || thumb !== undefined) {
    if (!key) bad('client_cert_key');
    if (!thumb) bad('client_cert_thumbprint');
    if (key!.length > 16384 || !/^[A-Za-z0-9+/]+={0,2}$/.test(key!)) bad('client_cert_key');
    privateKey = Buffer.from(key!, 'base64').toString('utf8');
    try {
      if (!privateKey.startsWith('-----BEGIN PRIVATE KEY-----')) bad('client_cert_key');
      createPrivateKey(privateKey);
    } catch {
      bad('client_cert_key');
    }
    thumbprint = thumb!.replace(/:/g, '').toLowerCase();
    if (!/^[a-f0-9]{64}$/.test(thumbprint)) bad('client_cert_thumbprint');
  } else if (
    !secret ||
    secret.length > 1024 ||
    !/^[\x21-\x7e](?:[\x20-\x7e]*[\x21-\x7e])?$/.test(secret)
  )
    bad('client_secret');
  const raw = header('mailbox');
  if (!raw?.trim()) bad('mailbox', 401);
  const mailbox = normalizeMailbox(raw!);
  if (!validMailbox(mailbox)) bad('mailbox');
  return { tenantId, clientId, mailbox, secret, privateKey, thumbprint };
}
