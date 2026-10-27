/* eslint-disable no-control-regex -- Deliberately reject control characters at the trust boundary. */
/**
 * Strict per-tool argument schemas for m365-shared-mailbox (WYREAI-325
 * security review round 2, R2-2).
 *
 * The recipient and mailbox-pin walkers in m365-shared-mailbox-policy.ts are
 * key DENYLISTS: they only see what they know how to look for. This module
 * turns that into an ALLOWLIST. Every classified tool has a strict schema
 * (allowed keys, types, sizes); the policy gate runs it first, so the walkers
 * only ever see schema-validated input.
 *
 * Rules, applied to every tool:
 *  - unknown keys are rejected at every level (no passthrough);
 *  - a string value that parses as a JSON object/array is rejected anywhere
 *    (a "stringified body" could smuggle recipients past the walkers);
 *  - ids and paths (any key ending in Id/ID, or named id/path/destinationId)
 *    must match an ALLOWLIST, `^[A-Za-z0-9_\-=+]{1,512}$` (R3-4): Graph ids
 *    are base64url-like, and the well-known folder names (inbox, drafts,
 *    sentitems, ...) are plain letters. Anything else, including `/`, `\`,
 *    `%`, `.`, `?`, `#`, `;`, `:`, any whitespace and every non-ASCII
 *    character (fullwidth or division slashes, NBSP, U+2028), is rejected:
 *    the sidecar substitutes these values into a Graph URL path.
 *  - `expand` / `$expand` must be absent or empty (R4-5, an allowlist of
 *    nothing): no tool needs a navigation property expanded (message and
 *    event reads return the full resource; folders and messages have their
 *    own list tools), and an expand of attachments would return their content
 *    inline. Attachment tools are hard-disabled.
 *  - `select` and `orderby` are token allowlists (R4-5): comma-separated
 *    property names (orderby: optionally followed by asc/desc); `select` must
 *    not name `attachments`. `*`, `%`, `$`, `&`, `(` and the like never pass.
 *  - free query strings (`filter`, `search`, `startDateTime`, `endDateTime`,
 *    `timezone`) must not contain `&`, `%`, `#`, a `$name` system-parameter
 *    token or a control character (R4-5). The sidecar encodes every query
 *    value exactly once (sidecar contract), so these are defence in depth
 *    against a second, injected query parameter. OData filter syntax
 *    (eq, and, quotes, parentheses, spaces, `/`, `:`) is unaffected. In
 *    `filter` only, `&`, `%`, `#` and `$word` are allowed inside an OData
 *    single-quoted literal (`'R&D'`, `''` escapes a quote); an unterminated
 *    literal is rejected (R5-4). `search` (KQL) stays strict.
 *  - recipient addresses are not trimmed: an address with leading or
 *    trailing whitespace, or any control or line-break character, is
 *    rejected (R3-5c). Every recipient and attendee address must also match
 *    the single-address grammar the send policy checks (SHARED_MAILBOX_ADDRESS_RE,
 *    case-insensitive), on every tool, drafts included (R4 info): no
 *    interior whitespace (any Unicode space, NBSP included) and no
 *    `Name <a@b.com>` display form.
 *  - a tool with no schema here is denied (the classified tool set is the
 *    source of truth; a test pins the two together).
 *
 * The mailbox is pinned by the connection, so there is deliberately NO userId
 * / mailbox / upn argument on any tool. Argument names follow softeria
 * ms-365-mcp-server's org-mode tools. Messages name the offending key, never
 * the value.
 */
import { z } from 'zod';

// ---------------------------------------------------------------------------
// Generic scans (run before the schema, independent of tool)
// ---------------------------------------------------------------------------

const MAX_DEPTH = 8;
const MAX_NODES = 5000;

export function isIdLikeKey(key: string): boolean {
  if (/(?:Id|ID)$/.test(key)) return true;
  const k = key.toLowerCase();
  return k === 'id' || k === 'path' || k === 'destinationid';
}

/** The only shape an identifier/path value may have (R3-4 allowlist). */
const SAFE_ID_RE = /^[A-Za-z0-9_\-=+]{1,512}$/;

/**
 * Graph well-known mail folder names, accepted wherever a folder id is
 * (mailFolderId, destinationId). Documentation and test fixture only: every
 * one of them already matches SAFE_ID_RE.
 */
export const GRAPH_WELL_KNOWN_FOLDER_NAMES: readonly string[] = [
  'archive',
  'clutter',
  'conflicts',
  'conversationhistory',
  'deleteditems',
  'drafts',
  'inbox',
  'junkemail',
  'localfailures',
  'msgfolderroot',
  'outbox',
  'recoverableitemsdeletions',
  'scheduled',
  'searchfolders',
  'sentitems',
  'serverfailures',
  'syncissues',
];

/** True when an identifier/path value is outside the allowlist (could change the Graph URL path). */
export function isUnsafeIdValue(value: string): boolean {
  return !SAFE_ID_RE.test(value);
}

/**
 * The single-address grammar (lower-case; test a lower-cased value). Shared
 * with the send policy (m365-shared-mailbox-policy.ts) so drafts and sends
 * accept exactly the same addresses.
 */
export const SHARED_MAILBOX_ADDRESS_RE =
  /^[a-z0-9._%+-]+@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z]{2,}$/;

/** True when `value` is a single plain address the send policy would accept (no trimming). */
export function isAcceptableRecipientAddress(value: string): boolean {
  return !hasUnsafeAddressChars(value) && SHARED_MAILBOX_ADDRESS_RE.test(value.toLowerCase());
}

/** True when an `expand` value is absent-equivalent: "" or an empty list. */
export function isEmptyExpand(value: unknown): boolean {
  return value === '' || (Array.isArray(value) && value.length === 0);
}

const SELECT_TOKEN_RE = /^[A-Za-z][A-Za-z0-9_]{0,63}$/;
const ORDERBY_TOKEN_RE = /^[A-Za-z][A-Za-z0-9_/]{0,127}(?: (?:asc|desc))?$/i;

function queryTokens(value: unknown): string[] | null {
  const items = Array.isArray(value) ? value : [value];
  if (!items.every((v) => typeof v === 'string')) return null;
  return (items as string[]).flatMap((v) => v.split(',')).map((t) => t.trim());
}

/** True when a `select` value is a list of plain property names, none of them `attachments`. */
export function isSafeSelect(value: unknown): boolean {
  const tokens = queryTokens(value);
  return (
    !!tokens && tokens.every((t) => SELECT_TOKEN_RE.test(t) && t.toLowerCase() !== 'attachments')
  );
}

/** True when an `orderby` value is a list of property paths, each optionally `asc`/`desc`. */
export function isSafeOrderby(value: unknown): boolean {
  const tokens = queryTokens(value);
  return !!tokens && tokens.every((t) => ORDERBY_TOKEN_RE.test(t));
}

/** Control and line-break characters: rejected anywhere in a free query string. */
const QUERY_CONTROL_RE = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/;
/** Could start a second query parameter or a pre-encoded value: `&`, `%`, `#`, `$name`. */
const QUERY_SEPARATOR_RE = /[&%#]|\$[A-Za-z]/;

/**
 * True when a free query string (search, dates, timezone) could carry
 * a second query parameter or a pre-encoded value: `&`, `%`, `#`, a
 * `$name` token (e.g. `$expand`), or a control/line-break character.
 */
export function hasQueryInjectionChars(value: string): boolean {
  return QUERY_CONTROL_RE.test(value) || QUERY_SEPARATOR_RE.test(value);
}

/**
 * The OData `filter` variant of hasQueryInjectionChars (WYREAI-325 R5-4).
 * Inside a single-quoted string literal (`''` is an escaped quote) `&`, `%`,
 * `#` and `$word` are ordinary text, so `contains(subject,'50%')` and
 * `subject eq 'R&D'` pass. Outside literals they are still rejected, and an
 * unterminated literal fails closed. Control characters are rejected
 * everywhere. Allowing them in literals relies on the sidecar encoding the
 * value exactly once (sidecar contract), which keeps them literal text.
 */
export function hasFilterInjectionChars(value: string): boolean {
  if (QUERY_CONTROL_RE.test(value)) return true;
  const outside = odataTextOutsideLiterals(value);
  return outside === null || QUERY_SEPARATOR_RE.test(outside);
}

/**
 * `value` with every OData single-quoted string literal replaced by `L`, or
 * null if a literal is unterminated. The placeholder is a letter so a `$`
 * directly before a literal still reads as a `$name` token (fail closed).
 */
function odataTextOutsideLiterals(value: string): string | null {
  let out = '';
  let i = 0;
  while (i < value.length) {
    if (value[i] !== "'") {
      out += value[i];
      i += 1;
      continue;
    }
    i += 1;
    for (;;) {
      if (i >= value.length) return null;
      if (value[i] === "'") {
        if (value[i + 1] !== "'") break;
        i += 2;
      } else {
        i += 1;
      }
    }
    i += 1;
    out += 'L';
  }
  return out;
}

/**
 * Free-text query keys and their check (with or without `$`). Only `filter`
 * is OData-tokenized: Graph `$search` takes KQL with double-quoted phrases,
 * not OData single-quoted literals, so it keeps the strict check.
 */
const FREE_QUERY_CHECKS: ReadonlyArray<readonly [string, (value: string) => boolean]> = [
  ['filter', hasFilterInjectionChars],
  ['search', hasQueryInjectionChars],
  ['startDateTime', hasQueryInjectionChars],
  ['endDateTime', hasQueryInjectionChars],
  ['timezone', hasQueryInjectionChars],
];

/** True when an `expand` value names attachments, in any case or list position. */
export function expandNamesAttachments(value: unknown): boolean {
  const items = Array.isArray(value) ? value : [value];
  return items.some((v) => typeof v === 'string' && /attachment/i.test(v));
}

/**
 * True when a recipient address carries leading/trailing whitespace or any
 * control or line-break character (C0, DEL, C1 incl. U+0085, U+2028, U+2029).
 * Addresses are never trimmed (R3-5c): what Conduit checks is what is sent.
 */
export function hasUnsafeAddressChars(value: string): boolean {
  return value !== value.trim() || /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/.test(value);
}

/** True when a string is itself a JSON object or array (a smuggled payload). */
export function isJsonContainerString(value: string): boolean {
  const s = value.trimStart();
  if (s[0] !== '{' && s[0] !== '[') return false;
  try {
    const parsed: unknown = JSON.parse(s);
    return typeof parsed === 'object' && parsed !== null;
  } catch {
    return false;
  }
}

/** Safe-to-print label for an argument path (keys are caller-controlled). */
function label(path: ReadonlyArray<string | number>): string {
  const text = path
    .map((p) =>
      String(p)
        .replace(/[^A-Za-z0-9_$-]/g, '?')
        .slice(0, 40)
    )
    .join('.');
  return text.slice(0, 80) || '(arguments)';
}

function scanGeneric(args: Record<string, unknown>): string | null {
  let nodes = 0;
  const walk = (
    value: unknown,
    path: Array<string | number>,
    idContext: boolean
  ): string | null => {
    if (path.length > MAX_DEPTH || ++nodes > MAX_NODES)
      return 'Tool arguments are too deeply nested.';
    if (typeof value === 'string') {
      if (idContext && isUnsafeIdValue(value)) {
        return `Argument "${label(path)}" must be a plain identifier (letters, digits, "-", "_", "=", "+").`;
      }
      if (isJsonContainerString(value)) {
        return `Argument "${label(path)}" must not contain a JSON-encoded object or array.`;
      }
      return null;
    }
    if (Array.isArray(value)) {
      for (let i = 0; i < value.length; i++) {
        const r = walk(value[i], [...path, i], idContext);
        if (r) return r;
      }
      return null;
    }
    if (typeof value === 'object' && value !== null) {
      for (const [k, v] of Object.entries(value)) {
        const r = walk(v, [...path, k], isIdLikeKey(k));
        if (r) return r;
      }
    }
    return null;
  };
  return walk(args, [], false);
}

// ---------------------------------------------------------------------------
// Building blocks
// ---------------------------------------------------------------------------

const str = (max = 2048) => z.string().max(max);
const idStr = z.string().min(1).max(512);
const bool = z.boolean();
const strList = (max = 100) => z.array(str(256)).max(max);

const emailAddress = z.strictObject({
  address: str(320).refine(isAcceptableRecipientAddress),
  name: str(256).optional(),
});
const recipient = z.strictObject({ emailAddress });
const recipients = z.array(recipient).max(500);

const bodyContent = z.strictObject({
  contentType: z.enum(['text', 'html', 'Text', 'HTML', 'Html']),
  content: z.string().max(4_000_000),
});

const dateTimeZone = z.strictObject({ dateTime: str(64), timeZone: str(128) });

const messageFlag = z.strictObject({
  flagStatus: z.enum(['notFlagged', 'complete', 'flagged']),
  startDateTime: dateTimeZone.optional(),
  dueDateTime: dateTimeZone.optional(),
  completedDateTime: dateTimeZone.optional(),
});

const importance = z.enum(['low', 'normal', 'high']);

/** Graph message resource. Attachments are intentionally not accepted. */
const message = z.strictObject({
  subject: str(1024).optional(),
  body: bodyContent.optional(),
  toRecipients: recipients.optional(),
  ccRecipients: recipients.optional(),
  bccRecipients: recipients.optional(),
  replyTo: recipients.optional(),
  from: recipient.optional(),
  sender: recipient.optional(),
  importance: importance.optional(),
  isRead: bool.optional(),
  isDeliveryReceiptRequested: bool.optional(),
  isReadReceiptRequested: bool.optional(),
  categories: strList().optional(),
  flag: messageFlag.optional(),
});

const address = z.strictObject({
  street: str(256).optional(),
  city: str(256).optional(),
  state: str(256).optional(),
  countryOrRegion: str(256).optional(),
  postalCode: str(64).optional(),
});

const location = z.strictObject({
  displayName: str(512).optional(),
  locationType: str(64).optional(),
  locationUri: str(2048).optional(),
  address: address.optional(),
  coordinates: z.strictObject({ latitude: z.number(), longitude: z.number() }).optional(),
});

const attendee = z.strictObject({
  emailAddress,
  type: z.enum(['required', 'optional', 'resource']).optional(),
});

const recurrence = z.strictObject({
  pattern: z.strictObject({
    type: str(64),
    interval: z.number().int().min(1).max(1000),
    month: z.number().int().min(0).max(12).optional(),
    dayOfMonth: z.number().int().min(0).max(31).optional(),
    daysOfWeek: strList(7).optional(),
    firstDayOfWeek: str(32).optional(),
    index: str(32).optional(),
  }),
  range: z.strictObject({
    type: str(64),
    startDate: str(64).optional(),
    endDate: str(64).optional(),
    recurrenceTimeZone: str(128).optional(),
    numberOfOccurrences: z.number().int().min(0).max(1000).optional(),
  }),
});

/** Graph event resource. */
const event = z.strictObject({
  subject: str(1024).optional(),
  body: bodyContent.optional(),
  start: dateTimeZone.optional(),
  end: dateTimeZone.optional(),
  location: location.optional(),
  locations: z.array(location).max(20).optional(),
  attendees: z.array(attendee).max(500).optional(),
  isOnlineMeeting: bool.optional(),
  onlineMeetingProvider: str(64).optional(),
  isAllDay: bool.optional(),
  isReminderOn: bool.optional(),
  reminderMinutesBeforeStart: z.number().int().min(0).max(525_600).optional(),
  showAs: str(32).optional(),
  sensitivity: str(32).optional(),
  importance: importance.optional(),
  categories: strList().optional(),
  allowNewTimeProposals: bool.optional(),
  hideAttendees: bool.optional(),
  responseRequested: bool.optional(),
  recurrence: recurrence.optional(),
  transactionId: idStr.optional(),
});

// OData-style query parameters. Accepted with or without the `$` prefix.
const QUERY_SHAPES = {
  top: z.number().int().min(0).max(1000),
  skip: z.number().int().min(0).max(100_000),
  filter: str(4096),
  select: z.union([str(1024), strList()]),
  orderby: z.union([str(512), strList(10)]),
  search: str(1024),
  count: bool,
  expand: z.union([str(512), strList(10)]),
} as const;

type QueryName = keyof typeof QUERY_SHAPES;

function queryShape(names: readonly QueryName[]): Record<string, z.ZodTypeAny> {
  const out: Record<string, z.ZodTypeAny> = {};
  for (const n of names) {
    out[n] = QUERY_SHAPES[n].optional();
    out[`$${n}`] = QUERY_SHAPES[n].optional();
  }
  return out;
}

const LIST_QUERY: readonly QueryName[] = [
  'top',
  'skip',
  'filter',
  'select',
  'orderby',
  'search',
  'count',
  'expand',
];
const READ_QUERY: readonly QueryName[] = ['select', 'expand'];

const timeWindow = { startDateTime: str(64), endDateTime: str(64), timezone: str(128).optional() };

function tool(shape: Record<string, z.ZodTypeAny>) {
  return z.strictObject(shape);
}

const list = (extra: Record<string, z.ZodTypeAny> = {}) =>
  tool({ ...extra, ...queryShape(LIST_QUERY), fetchAllPages: bool.optional() });
const read = (extra: Record<string, z.ZodTypeAny> = {}) =>
  tool({ ...extra, ...queryShape(READ_QUERY) });
const withBody = (ids: Record<string, z.ZodTypeAny>, body: z.ZodTypeAny) => tool({ ...ids, body });
const idOnly = (...names: string[]) => tool(Object.fromEntries(names.map((n) => [n, idStr])));

const commentBody = z.strictObject({
  comment: str(100_000).optional(),
  message: message.optional(),
});
const forwardBody = z.strictObject({
  comment: str(100_000).optional(),
  toRecipients: recipients,
  message: message.optional(),
});
const responseBody = z.strictObject({
  comment: str(100_000).optional(),
  sendResponse: bool.optional(),
});
const destinationBody = z.strictObject({ destinationId: idStr });

// ---------------------------------------------------------------------------
// One strict schema per classified tool
// ---------------------------------------------------------------------------

export const SHARED_MAILBOX_TOOL_SCHEMAS: Readonly<Record<string, z.ZodTypeAny>> = {
  // reads
  'get-calendar-event': read({ eventId: idStr, timezone: str(128).optional() }),
  'get-specific-calendar-event': read({
    calendarId: idStr,
    eventId: idStr,
    timezone: str(128).optional(),
  }),
  'get-calendar-view': list(timeWindow),
  'get-specific-calendar-view': list({ calendarId: idStr, ...timeWindow }),
  'get-mail-message': read({ messageId: idStr }),
  'list-calendar-event-instances': list({ eventId: idStr, ...timeWindow }),
  'list-calendar-events': list({ timezone: str(128).optional() }),
  'list-specific-calendar-events': list({ calendarId: idStr, timezone: str(128).optional() }),
  'list-calendars': list(),
  'list-mail-child-folders': list({ mailFolderId: idStr }),
  'list-mail-folder-messages': list({ mailFolderId: idStr }),
  'list-mail-folders': list(),
  'list-mail-messages': list(),

  // triage and drafts (write)
  'copy-mail-message': withBody({ messageId: idStr }, destinationBody),
  'move-mail-message': withBody({ messageId: idStr }, destinationBody),
  'update-mail-message': withBody({ messageId: idStr }, message),
  'create-draft-email': tool({ body: message }),
  'create-reply-draft': withBody({ messageId: idStr }, commentBody),
  'create-reply-all-draft': withBody({ messageId: idStr }, commentBody),
  'create-forward-draft': withBody(
    { messageId: idStr },
    z.strictObject({ ...forwardBody.shape, toRecipients: recipients.optional() })
  ),

  // admin: delete and outbound
  'delete-mail-message': idOnly('messageId'),
  'send-draft-message': idOnly('messageId'),
  'send-mail': tool({ body: z.strictObject({ message, saveToSentItems: bool.optional() }) }),
  'reply-mail-message': withBody({ messageId: idStr }, commentBody),
  'reply-all-mail-message': withBody({ messageId: idStr }, commentBody),
  'forward-mail-message': withBody({ messageId: idStr }, forwardBody),
  'create-calendar-event': tool({ body: event }),
  'create-specific-calendar-event': withBody({ calendarId: idStr }, event),
  'update-calendar-event': withBody({ eventId: idStr }, event),
  'update-specific-calendar-event': withBody({ calendarId: idStr, eventId: idStr }, event),
  'delete-calendar-event': idOnly('eventId'),
  'delete-specific-calendar-event': idOnly('calendarId', 'eventId'),
  'accept-calendar-event': withBody({ eventId: idStr }, responseBody),
  'decline-calendar-event': withBody({ eventId: idStr }, responseBody),
  'tentatively-accept-calendar-event': withBody({ eventId: idStr }, responseBody),
  'cancel-calendar-event': withBody(
    { eventId: idStr },
    z.strictObject({ comment: str(100_000).optional() })
  ),
  'forward-calendar-event': withBody({ eventId: idStr }, forwardBody),
};

/** Every tool this module has a schema for (contract + completeness tests). */
export function sharedMailboxSchemaToolNames(): string[] {
  return Object.keys(SHARED_MAILBOX_TOOL_SCHEMAS).sort();
}

/**
 * Validate a call's arguments against the tool's strict schema. Returns a
 * denial message naming the offending key (never a value), or null when the
 * arguments are well-formed. A tool without a schema is denied.
 */
export function validateSharedMailboxToolArguments(
  toolName: string,
  toolArguments: Record<string, unknown> | undefined
): string | null {
  if (!Object.prototype.hasOwnProperty.call(SHARED_MAILBOX_TOOL_SCHEMAS, toolName)) {
    return 'This tool is not available for this connection.';
  }
  const args = toolArguments ?? {};
  if (typeof args !== 'object' || args === null || Array.isArray(args)) {
    return 'Tool arguments must be an object.';
  }
  const parsed = SHARED_MAILBOX_TOOL_SCHEMAS[toolName]!.safeParse(args);
  // Prefer reporting an unknown key: it is the more useful and more telling
  // signal (e.g. `userId: "ceo@..."` would otherwise be reported by the id
  // allowlist). Only the key is named, never the value.
  const unknown = parsed.success
    ? undefined
    : parsed.error.issues.find((i) => i.code === 'unrecognized_keys');
  if (unknown && unknown.code === 'unrecognized_keys') {
    const key = unknown.keys[0] ?? '';
    return `Argument "${label([...unknown.path, key])}" is not accepted by "${toolName}".`;
  }

  const scan = scanGeneric(args);
  if (scan) return scan;
  for (const key of ['expand', '$expand']) {
    if (!Object.prototype.hasOwnProperty.call(args, key)) continue;
    if (expandNamesAttachments(args[key])) {
      return `Argument "${key}" must not expand attachments; attachment content is not available for this connection.`;
    }
    if (!isEmptyExpand(args[key])) {
      return `Argument "${key}" is not supported for this connection; leave it empty.`;
    }
  }
  for (const key of ['select', '$select']) {
    if (Object.prototype.hasOwnProperty.call(args, key) && !isSafeSelect(args[key])) {
      return `Argument "${key}" must be a comma-separated list of property names (attachments cannot be selected).`;
    }
  }
  for (const key of ['orderby', '$orderby']) {
    if (Object.prototype.hasOwnProperty.call(args, key) && !isSafeOrderby(args[key])) {
      return `Argument "${key}" must be a comma-separated list of property names, each optionally followed by asc or desc.`;
    }
  }
  for (const [name, isUnsafe] of FREE_QUERY_CHECKS) {
    for (const key of [name, `$${name}`]) {
      const v = args[key];
      if (Object.prototype.hasOwnProperty.call(args, key) && typeof v === 'string' && isUnsafe(v)) {
        return name === 'filter'
          ? `Argument "${key}" must not contain "&", "%", "#" or a "$" query option outside a quoted string, an unterminated quote, or a control character.`
          : `Argument "${key}" must not contain "&", "%", "#", a "$" query option or a control character.`;
      }
    }
  }

  if (parsed.success) return null;
  return `Argument "${label(parsed.error.issues[0]!.path)}" is missing or invalid for "${toolName}".`;
}
