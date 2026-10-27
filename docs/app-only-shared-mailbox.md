# App-only shared-mailbox mode (WYRE-AI fork)

This fork adds one optional mode to Softeria's ms-365-mcp-server: an HTTP MCP server that works
on **one Exchange Online shared mailbox per request**, using an Entra **app-only**
(client-credentials) token. It is built for the WYRE-AI Conduit gateway, which runs it as the
`m365-shared-mailbox-mcp` sidecar. It is **off by default**; nothing about the upstream server
changes unless you turn it on.

Access to the mailbox is granted with **Exchange Online RBAC for Applications**, scoped to that one
mailbox. Do **not** grant tenant-wide Microsoft Graph `Mail.*`, `Calendars.*` or
`MailboxSettings.*` application permissions: the server refuses to run with them (see
[Tenant-wide consent check](#tenant-wide-consent-check)).

## Turning it on

- Docker: `Dockerfile.app-only` builds the image. It runs `node dist/app-only/main.js` with
  `MS365_MCP_APP_ONLY_SHARED_MAILBOX=1`.
- CLI: `ms-365-mcp-server --app-only-shared-mailbox` (or set `MS365_MCP_APP_ONLY_SHARED_MAILBOX=1`).
  No other CLI flag is accepted in this mode, and startup fails if upstream auth settings such as
  extra scopes, trust-proxy auth, OBO or OAuth are present in the environment.

| Variable                           | Default            | Meaning                                                                                                                    |
| ---------------------------------- | ------------------ | -------------------------------------------------------------------------------------------------------------------------- |
| `CONDUIT_S2S_SECRET`               | required           | Gateway service-to-service secret. Empty or a known placeholder refuses startup.                                           |
| `MCP_ALLOW_INSECURE_DEV`           | unset              | `1` allows an empty secret for local development only, and forces `127.0.0.1`.                                             |
| `MCP_HTTP_HOST` / `MCP_HTTP_PORT`  | `0.0.0.0` / `8080` | Listen address.                                                                                                            |
| `M365_APP_ONLY_MAX_TIER`           | `write`            | `read`, `write` or `admin`. Tools above the tier are not listed and cannot be called. `READ_ONLY_MODE=true` forces `read`. |
| `M365_APP_ONLY_PINNED_MAILBOX`     | unset              | Optional: the only mailbox this instance will serve.                                                                       |
| `M365_APP_ONLY_BIND_MAILBOX`       | `1`                | Bind each credential to the first mailbox it is used with (see [Mailbox lock](#mailbox-lock)).                             |
| `M365_APP_ONLY_BIND_TTL_SECONDS`   | `3600`             | Idle lifetime of that binding.                                                                                             |
| `M365_APP_ONLY_MAX_BODY_CHARS`     | `20000`            | Per-body cap after HTML-to-text.                                                                                           |
| `M365_APP_ONLY_MAX_RESPONSE_BYTES` | `4000000`          | Cap on any Graph response read.                                                                                            |
| `M365_APP_ONLY_MAX_PAGES`          | `10`               | Page cap for `fetchAllPages`.                                                                                              |
| `M365_APP_ONLY_MAX_REQUEST_BYTES`  | `5000000`          | Cap on an incoming MCP request body.                                                                                       |
| `M365_APP_ONLY_LOG_LEVEL`          | `info`             | JSON log level.                                                                                                            |

## HTTP surface

- `GET /health` and `GET /healthz`: `200 {"status":"ok"}`. No authentication, no token request,
  no Graph call.
- `POST /mcp`: MCP over Streamable HTTP, stateless, JSON responses.
- Everything else: `404`. There are no OAuth, `.well-known`, attachment or download routes.

Every request except health must carry `x-gateway-s2s: t=<unix seconds>,v1=<hex HMAC-SHA256(secret, "t=<t>")>`
(300 s skew), the same scheme as the other WYRE-AI fleet sidecars. Without it the answer is `401`
before anything else happens.

### Per-request credentials

| Header                                                     | Value                                                                              |
| ---------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `X-M365-Tenant-Id`                                         | Directory (tenant) ID, GUID                                                        |
| `X-M365-Client-Id`                                         | Application (client) ID, GUID                                                      |
| `X-M365-Client-Secret`                                     | Client secret, **or** the certificate pair below                                   |
| `X-M365-Client-Cert-Key` + `X-M365-Client-Cert-Thumbprint` | Base64 of an unencrypted PKCS#8 PEM private key, and the cert's SHA-256 thumbprint |
| `X-M365-Mailbox`                                           | The shared mailbox's primary SMTP address                                          |

The certificate pair promised to Conduit is exactly `X-M365-Client-Cert-Key` and
`X-M365-Client-Cert-Thumbprint`. Shorter names such as `X-M365-Cert-Key` and
`X-M365-Cert-Thumbprint` are ignored. Conduit currently sends only the client-secret form.

Sending a secret and a certificate together is `400 bad_credentials:client_secret_and_certificate`.
Other malformed values are `400 bad_credentials:<field>`; a missing mailbox header is
`401 bad_credentials:mailbox`. Credentials are never read from the environment in this mode.

Tokens are requested with MSAL `acquireTokenByClientCredential` for
`https://graph.microsoft.com/.default` and kept in memory only, keyed by a keyed hash of tenant,
client and credential.

## Mailbox lock

The mailbox comes **only** from `X-M365-Mailbox` (trimmed, lower-cased, Unicode NFC). Every Graph
URL is `https://graph.microsoft.com/v1.0/users/{mailbox}/...`, with the mailbox and every path
parameter passed through `encodeURIComponent`. There is no `/me`, no default mailbox and no
`/users` listing, and no tool accepts `userId`, `mailbox`, `upn` or `userPrincipalName`.

`mailbox_mismatch` (HTTP 403, before any Graph call) is returned when:

- `M365_APP_ONLY_PINNED_MAILBOX` is set and the header differs;
- the same credential was used with a different mailbox within the binding TTL;
- a `fetchAllPages` next link points anywhere other than the same mailbox.

Exchange RBAC for Applications is the real boundary; the lock is defense in depth.

## Tools

Only these tools exist. Everything else, including upstream's `graph-batch`, attachment and
download tools, mail rules, mailbox settings and user or directory tools, is absent from
`tools/list`. A direct call returns JSON-RPC `-32601 tool_disabled_app_only`. There are no MCP
resources or prompts.

- **read** (13): `list-mail-messages`, `list-mail-folders`, `list-mail-child-folders`,
  `list-mail-folder-messages`, `get-mail-message`, `list-calendars`, `list-calendar-events`,
  `list-specific-calendar-events`, `get-calendar-event`, `get-specific-calendar-event`,
  `get-calendar-view`, `get-specific-calendar-view`, `list-calendar-event-instances`.
- **write** (7, drafts and triage): `create-draft-email`, `create-reply-draft`,
  `create-reply-all-draft`, `create-forward-draft`, `update-mail-message`, `move-mail-message`,
  `copy-mail-message`.
- **admin** (17, outbound and deletes): `send-mail`, `send-draft-message`, `reply-mail-message`,
  `reply-all-mail-message`, `forward-mail-message`, `delete-mail-message`, `create-calendar-event`,
  `create-specific-calendar-event`, `update-calendar-event`, `update-specific-calendar-event`,
  `delete-calendar-event`, `delete-specific-calendar-event`, `accept-calendar-event`,
  `decline-calendar-event`, `tentatively-accept-calendar-event`, `cancel-calendar-event`,
  `forward-calendar-event`.

Arguments follow the contract from WYRE-AI/conduit #2264 @4c5f375c. Strict schemas reject
unknown keys at every level. Arguments are parsed once; JSON-encoded objects or arrays in string
fields are rejected. A rejected call returns JSON-RPC `-32602 invalid_arguments`, with no token
request or Graph call.

- Path IDs must match the allowlist `^[A-Za-z0-9_\-=+]{1,512}$` and are also encoded with
  `encodeURIComponent`. Non-ID strings never enter the path.
- `expand` must be absent, `""` or `[]`. `select` accepts comma-separated or array entries of plain
  property names (`^[A-Za-z][A-Za-z0-9_]{0,63}$`), never `attachments`.
- Each `orderby` token matches `^[A-Za-z][A-Za-z0-9_/]{0,127}( (asc|desc))?$`
  (case-insensitive).
- `filter` allows `&`, `%`, `#` and `$word` only inside single-quoted OData literals, with `''`
  escaping a quote. Controls and line breaks are always rejected. `search` keeps the strict ban on
  those characters even inside quotes (known limitation; use `filter` with `contains` instead).
- Recipient and attendee addresses must match `ADDRESS_RE` exactly (case-insensitive):
  `^[a-z0-9._%+-]+@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z]{2,}$`.
  Addresses are never trimmed or modified; whitespace and display forms such as `Name <a@b.com>`
  are rejected.
- Query values are encoded exactly once through `URLSearchParams.set`, without pre-decoding or
  manual query-pair concatenation.

The fixture test iterates `tools`/`validExample`, `mustReject`, `mustAccept` and `hardDisabledTools`.

R6-2 (encode-once filter literal pilot gate) checks the `filter-literal-encode-once` fixture
end-to-end: Graph receives exactly one `$filter`, no `$expand`, and the unchanged literal
after one decode, with no double encoding.

**Pilot runbook:** `create-draft-email`, `create-forward-draft` and `update-mail-message` accept any
well-formed recipient. Conduit denies `send-draft-message`. A human must review every draft,
including its recipients, in Outlook before sending.

Message and event bodies are returned as text, capped, and wrapped in
`<untrusted-content nonce="...">` markers: they are external content and must be treated as data.

## Errors

| Code                                             | When                                                              |
| ------------------------------------------------ | ----------------------------------------------------------------- |
| `401 unauthorized`                               | Missing or invalid S2S header                                     |
| `400 bad_credentials:<field>`                    | Malformed credential header (`401` for a missing mailbox header)  |
| `401 auth_failed:secret_expired`                 | `AADSTS7000222`                                                   |
| `401 auth_failed:invalid_client`                 | `AADSTS7000215` (wrong secret), certificate errors                |
| `401 auth_failed:tenant_or_app`                  | `AADSTS90002` / `AADSTS700016`, or a token for another tenant/app |
| `403 tenant_wide_consent_detected`               | See below                                                         |
| `403 mailbox_mismatch`                           | See [Mailbox lock](#mailbox-lock)                                 |
| `graph_forbidden:check_exchange_rbac_scope`      | Graph 403: the RBAC scope does not cover this mailbox or action   |
| `graph_throttled:retry_after=<s>`                | Graph 429/503. Retried once when the wait is 10 s or less         |
| `graph_not_found`, `graph_error:<status>:<code>` | Other Graph failures                                              |

Request-level errors are JSON-RPC errors with code `-32001`; Graph errors are tool results with
`isError: true`. Logs never contain secrets, keys, tokens, argument values or message bodies.

## Tenant-wide consent check

Each token's `roles` claim is decoded. If it holds any `Mail.*`, `Calendars.*` or
`MailboxSettings.*` application role, the app can reach **every** mailbox in the tenant, and the
server refuses with HTTP 403, JSON-RPC `-32001 tenant_wide_consent_detected`, listing the roles in
`data.roles` and the fix in `data.fix`. Remove those Graph application permissions (and their admin
consent) from the app registration and use the RBAC setup below.

## Customer setup: Exchange Online RBAC for Applications

Run these once per shared mailbox, as an Exchange administrator, in Exchange Online PowerShell
(`Connect-ExchangeOnline`). Replace the placeholders.

1. **Register an app** in Entra ID (single tenant). Add a client secret or, preferably, upload a
   certificate. **Do not** add Microsoft Graph `Mail.*`, `Calendars.*` or `MailboxSettings.*`
   application permissions and do not grant admin consent for them. Note the Application (client)
   ID and the Object ID of its **Enterprise application** (service principal).
2. **Create the Exchange service principal pointer**:

   ```powershell
   New-ServicePrincipal -AppId <client-id> -ObjectId <enterprise-app-object-id> -DisplayName "Conduit shared mailbox"
   ```

3. **Create a management scope for the one mailbox**:

   ```powershell
   New-ManagementScope -Name "Conduit - support mailbox" `
     -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'support@contoso.com'"
   ```

4. **Assign only the roles you need, scoped to it.** Drafting and triage need
   `Application Mail.ReadWrite`; sending needs `Application Mail.Send`; calendar tools need
   `Application Calendars.ReadWrite` (or `Application Calendars.Read`):

   ```powershell
   New-ManagementRoleAssignment -App <client-id> -Role "Application Mail.ReadWrite" `
     -CustomResourceScope "Conduit - support mailbox"
   New-ManagementRoleAssignment -App <client-id> -Role "Application Mail.Send" `
     -CustomResourceScope "Conduit - support mailbox"
   New-ManagementRoleAssignment -App <client-id> -Role "Application Calendars.ReadWrite" `
     -CustomResourceScope "Conduit - support mailbox"
   ```

5. **Check it**. The first command should be `InScope: True`, the second `False`:

   ```powershell
   Test-ServicePrincipalAuthorization -Identity <client-id> -Resource support@contoso.com
   Test-ServicePrincipalAuthorization -Identity <client-id> -Resource someone.else@contoso.com
   ```

Notes:

- Permissions are a **union**: any tenant-wide Graph grant on the same app overrides the scope,
  which is why the server refuses tokens that carry one.
- Changes can take **30 minutes to 2 hours** to apply. Until then Graph may return 403
  (`graph_forbidden:check_exchange_rbac_scope`).
- Only the global Microsoft cloud is supported in this mode.

## Threat model (summary)

- **Cross-mailbox access**: blocked by the RBAC scope, the header-only mailbox, strict schemas with
  no mailbox arguments, path-parameter validation plus encoding, and the tenant-wide consent check.
- **Prompt injection via mail**: bodies are wrapped and capped; outbound tools are admin tier and off
  by default; Conduit also enforces recipient allowlists and quotas.
- **Gateway impersonation by a sibling container**: per-vendor S2S HMAC. The S2S proof does not bind
  the headers or body, so ingress must stay internal to the container environment.
- **Credential leakage**: credentials live only in request headers and in memory; logs redact them.
- **Upstream surface**: OAuth, OBO, trust-proxy auth, attachments, downloads, batch and extra
  scopes are never loaded by the app-only entry.
