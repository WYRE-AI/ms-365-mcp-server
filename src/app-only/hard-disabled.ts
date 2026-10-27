export const HARD_DISABLED_TOOLS = [
  'graph-batch',
  'send-batch',
  'create-mail-rule',
  'list-mail-rules',
  'get-mail-rule',
  'update-mail-rule',
  'delete-mail-rule',
  'get-mailbox-settings',
  'update-mailbox-settings',
  'list-mail-attachments',
  'get-mail-attachment',
  'add-mail-attachment',
  'download-bytes',
  'get-download-url',
  'download-attachment',
  'list-users',
  'get-user',
  'get-current-user',
  'list-directory-objects',
] as const;
export const hardDisabled = (name: string): boolean =>
  (HARD_DISABLED_TOOLS as readonly string[]).includes(name) ||
  /-attachment|-mail-rule|-mailbox-settings/.test(name);
