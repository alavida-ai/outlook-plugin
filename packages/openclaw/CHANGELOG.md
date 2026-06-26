# @alavida-ai/outlook-plugin-openclaw

## 0.4.2

### Patch Changes

- 2f07a6d: Fix `outlook_auth_login`'s hint so the sign-in link actually reaches the user.
  The previous hint said the link "has been sent", so a well-behaved agent
  concluded there was nothing to do and ended its turn — leaving the stashed URL
  undelivered until the user manually prompted it. The link is only delivered when
  the agent calls the `message` tool (the `message_sending` hook rewrites that
  outgoing message to carry the verbatim link). The hint now explicitly instructs
  the agent to call `message` to deliver the link, and the skill auth reference is
  updated to match.

## 0.4.1

### Patch Changes

- 4c032c1: Build against openclaw `2026.6.10` (latest stable; was `2026.5.12`). The newer
  plugin SDK surfaced a declaration-emit portability error (TS2742) on the default
  export, fixed by annotating it with the public `OpenClawPluginDefinition` type
  instead of the SDK-internal inferred return type. No runtime behavior change.
  The `peerDependencies` / `minGatewayVersion` floor is intentionally left at
  `>=2026.3.24-beta.2` so the plugin keeps loading on older gateways.
- fe9e76b: Stop logging the signed-in user's email (`upn`) in the auth-callback success
  line. The handler now records `sign-in complete for agent=<id>` only, so no
  personal data lands in gateway server logs. The token exchange still returns
  the `upn` internally; it is simply no longer written to stderr.

## 0.4.0

### Minor Changes

- 174c133: Deliver the `outlook_auth_login` sign-in URL out-of-band so it never passes
  through the agent. The tool now stashes the URL keyed by session and returns a
  sanitized envelope (`delivery: "channel"`, no `authUrl`); a `message_sending`
  hook rewrites the agent's next outbound reply in that session to carry the
  verbatim link. Because the agent never holds the URL, a prompt-injected agent
  can't swap in a phishing link. When there is no channel session to deliver to,
  the tool falls back to returning the URL inline (`delivery: "inline"`).

## 0.3.0

### Minor Changes

- fc1321c: Auth cleanup: remove the device-code flow and drop CLI multi-account support.
  - **core**: remove `loginDeviceCode` / `loginDeviceCodeInBackground` (and their
    types); the `LoginResult` type moves to its own module and is still exported.
  - **openclaw**: `outlook_auth_login` is now browser-only (Authorization Code +
    PKCE) and **requires** `oauthRedirectUri` — it returns a clear error if unset.
    Device-code is gone (it's blocked by the Conditional Access baselines we
    target, and the localhost interactive flow can't run on a headless gateway).
  - **cli**: single-account. `--account` / `OUTLOOK_ACCOUNT` are removed and
    `outlook auth login` now clears any cached account before signing in, so the
    CLI cache always holds exactly one identity.

### Patch Changes

- Updated dependencies [fc1321c]
  - @alavida-ai/outlook-core@0.3.0

## 0.2.1

### Patch Changes

- 721d4f8: Fix browser-flow auth docs: register the callback redirect URI under **Mobile
  and desktop applications** (public-client/Native type), not **Web**. The plugin
  is a public client (PKCE, no secret); a Web redirect makes Entra treat it as
  confidential and fail the token exchange with AADSTS7000218. Adds a note
  explaining why, and that the "Allow public client flows" toggle does not apply
  to the auth-code flow.

## 0.2.0

### Minor Changes

- 655cade: Add the Authorization Code + PKCE sign-in flow alongside device-code.
  - **core**: new `buildAuthCodeUrl` / `exchangeAuthCode` helpers (PKCE verifier,
    CSRF `state`, ID-token `nonce`, nonce verification on exchange) and a
    `loginInteractive` helper for the localhost-loopback interactive flow.
  - **openclaw**: `outlook_auth_login` returns a browser sign-in URL when the new
    `oauthRedirectUri` plugin config is set, and a new `/outlook/auth-callback`
    HTTP route (plugin-scoped, exact-match) redeems the code into the initiating
    agent's token cache. Pending flows are single-use with a 10-minute TTL. When
    `oauthRedirectUri` is unset, the existing device-code flow is unchanged.
  - **cli**: `outlook auth login` now uses the interactive browser flow
    (`acquireTokenInteractive`) instead of device-code, so it works on tenants
    whose Conditional Access blocks device-code sign-in.

### Patch Changes

- Updated dependencies [655cade]
  - @alavida-ai/outlook-core@0.2.0

## 0.1.1

### Patch Changes

- c2a4bdd: Fix every tool showing as `(anonymous)` in `openclaw plugins inspect outlook --runtime`.

  The plugin registers each tool as a **factory** (so the per-agent token cache can be baked into the config at agent-setup time, not module load time). When you register a factory, openclaw cannot introspect the tool's name without invoking the factory — and the inspector deliberately doesn't invoke factories at inspect time. The SDK provides a separate `opts.name` argument on `api.registerTool` for exactly this case; we were not passing it.

  `registerTool` in `packages/openclaw/src/register.ts` now passes `{ name: descriptor.name }` as the second argument. After updating, `openclaw plugins inspect outlook --runtime` will display the real tool names (`outlook_mail_list`, `outlook_auth_login`, etc.) instead of 17 lines of `(anonymous)`.

  This is a cosmetic fix — tools were callable from agents before this change. Worth shipping because the inspector display is the primary way operators sanity-check that a plugin loaded correctly, and because the `optional`/manifest-alignment flow depends on the same name hint.

## 0.1.0

### Minor Changes

- bfa8f53: Per-agent token cache + plugin-side auth tools.

  ## What changed

  On a multi-agent OpenClaw gateway, the Outlook plugin now stores tokens at
  **`<agentDir>/outlook-tokens.json`** (per-agent isolation) instead of one
  shared `~/.outlook-plugin/tokens.json`. Two agents on the same gateway can
  hold tokens for two different Microsoft accounts without cross-contamination.

  The CLI's cache is unchanged (`~/.outlook-plugin/tokens.json`) and is now
  **always isolated from agent caches** — even with a single agent on the host.

  ### `@alavida-ai/outlook-core` (minor)
  - New export: `loginDeviceCodeInBackground({ app, cache })` — non-blocking
    variant of `loginDeviceCode`. Returns `{ verificationUrl, userCode, expiresAt, completion }`
    as soon as Microsoft emits the device code (~200 ms); the `completion`
    promise resolves later when MSAL polling finishes. Used by the new
    `auth_login` plugin tool to avoid 15-minute-blocking tool calls.

  ### `@alavida-ai/outlook-plugin-openclaw` (minor)
  - **`registerTool` now passes a factory to OpenClaw** so each agent's tool
    invocation receives that agent's trusted `OpenClawPluginToolContext`
    (`agentId`, `agentDir`, `sessionKey`, …). The factory closes over the
    context and bakes a per-agent cache path into the returned tool.
  - **`PluginConfig` gained `agentId` and `agentDir` fields**, populated from
    the factory context at tool-call time.
  - **`getClient` is now memoised per-agent** (`Map<key, OutlookClient>` keyed
    on `clientId|tenantId|resolvedCachePath`) instead of a single-slot cache.
  - **Cache path resolution** (in priority order):
    1. `config.tokenCachePath` (explicit operator override)
    2. `OUTLOOK_TOKEN_CACHE` env
    3. `<agentDir>/outlook-tokens.json` (per-agent default)
    4. `~/.outlook-plugin/tokens.json` (standalone fallback)
  - **Three new tools:**
    - `auth_login` — fire-and-forget device-code login. Returns URL+code
      immediately for the agent to surface to the human; completes in
      background.
    - `auth_status` — reports `authenticated` / `not_authenticated` with a
      structured `reason`. Never throws auth errors itself.
    - `auth_logout` — wipes this agent's cache file.
  - Skill bundle (`skills/outlook/references/auth.md`) updated to describe the
    per-agent flow and the new tools.

  ### `@alavida-ai/outlook-cli` (patch)
  - No functional change. Bumped because it depends on
    `@alavida-ai/outlook-core` and `@alavida-ai/outlook-plugin-openclaw` via
    `workspace:*`.

  ## Migration

  Nothing is deployed yet; no migration steps. Operators who've already run
  `outlook auth login` on the host still have a working CLI cache at
  `~/.outlook-plugin/tokens.json` — only the plugin moves to per-agent paths.
  To authenticate the plugin for the first time, an agent calls
  `outlook.auth_login` inside its session.

- 2bd3ebc: Sun Global pilot hardening: scope downgrade, 30-min mail safety window, cloud.microsoft inbox URLs, plain-text body defaults, and `outlook_` prefix on all openclaw tool names.

  ## Breaking changes
  - **Microsoft Graph scope set is narrower.** `Calendars.ReadWrite` and `Calendars.ReadWrite.Shared` are now `Calendars.Read` and `Calendars.Read.Shared`. `Contacts.ReadWrite` has been removed entirely. Existing tokens cached under the old scope set must be refreshed — operators need to run `outlook auth logout` once per surface (CLI cache, each agent cache) and re-auth. Tenant admins will see the updated (read-only calendar, no contacts) consent screen on the next sign-in.
  - **10 CLI commands removed**: `mail {move,delete,mark,flag,importance}`, `calendar {create,update,delete,respond}`, `contacts list`.
  - **10 openclaw tools removed**: same surface as the CLI deletions.
  - **All openclaw tool names prefixed with `outlook_`** (e.g. `mail_list` → `outlook_mail_list`, `whoami` → `outlook_whoami`, `auth_login` → `outlook_auth_login`). Plugin manifest (`openclaw.plugin.json` `contracts.tools[]`) updated. Any gateway config or agent skill that references the old names must be updated.
  - **`outlook calendar show` defaults to plain-text body** instead of HTML. `--html` flag opts back into the raw HTML. Same for the openclaw `outlook_calendar_show` tool (defaults to `preferText: true`).
  - **`outlook_mail_read` openclaw tool defaults to plain-text body** instead of HTML. Pass `preferText: false` to opt back into raw HTML. CLI `outlook mail read` behavior unchanged.
  - **30-minute safety window applies to inbound mail.** Non-draft messages with `receivedDateTime` newer than `now - 30min` are filtered from `outlook_mail_list` / `outlook_mail_search`, and `outlook_mail_read` / `outlook_mail_reply` / `outlook_mail_forward` / `outlook_mail_list_attachments` / `outlook_mail_download_attachment` refuse them with a new `MailQuarantinedError` (openclaw error envelope: `{ error: 'mail_quarantined', availableAt, receivedDateTime }`). Drafts (`isDraft: true`) are exempt. This protects one-time passwords and 2FA codes from leaking into agent context.

  ## New features
  - **`inboxLink` field** on `outlook_mail_list` / `_read` / `_search` tool outputs and CLI JSON output. Builds `https://outlook.cloud.microsoft/mail/inbox/id/<id>` by translating REST ids to `restImmutableEntryId` via `POST /me/translateExchangeIds` (one batch round trip per command). The legacy `webLink` (Graph's OWA single-item URL) is still present alongside it. Prefer `inboxLink` when surfacing a message to the user.
  - **New `MailResource.inboxLinks(restIds[])` core export** — public batch translator returning `{ restId: url | null }`.
  - **New `MailQuarantinedError` core export** — extends `CoreError` with `messageId`, `receivedDateTime`, `availableAt` fields.
  - **New `inboxLinkFromId(id)` core helper** — pure builder for the cloud URL once you have an immutable id.
  - **`isDraft` field** on `MessageSummary` (was already on `Message` from Graph; now selected + surfaced).
  - **CLI: dedicated `Mail blocked: …` formatter** for `MailQuarantinedError`, prints `Received` + `Available at` timestamps to stderr.
  - **CLI: row-level `open:` URL in `mail list`** pretty rendering.

  ## Skills
  - **`packages/openclaw/skills/outlook/` rewritten** — all tool references prefixed, removed-tool mentions purged, new sections covering the safety window, the `inboxLink` UX, plain-text defaults, and the read-only calendar posture.
  - **`packages/cli/skills/outlook-cli/` is new** — parallel skill targeted at humans using the `outlook` CLI from Claude Code conversations. Bash + jq examples instead of TypeScript tool calls. Same safety posture, same scope constraints.

  ## `@alavida-ai/outlook-core` (minor)
  - New: `MailQuarantinedError`, `MAIL_QUARANTINE_MINUTES` (= 30), `mailQuarantineCutoffIso()`, `inboxLinkFromId()`, `MailResource.inboxLinks()`, `GetEventOptions`.
  - Changed: `OUTLOOK_SCOPES` is now `['Mail.ReadWrite', 'Calendars.Read', 'Calendars.Read.Shared', 'User.Read']`.
  - Changed: `MessageSummary` includes `isDraft`.
  - Changed: `CalendarResource.get(id, { preferText? })` accepts the new option.
  - Removed: `CalendarResource.create / update / delete / respond` methods and their input/output types (`CreateEventInput`, `UpdateEventInput`, `RespondInput`, `RECURRENCE_PRESETS`, `RecurrencePreset`, `ATTENDEE_RESPONSES`, `AttendeeResponse`).

  ## `@alavida-ai/outlook-plugin-openclaw` (minor)
  - All 17 tools prefixed with `outlook_`.
  - `outlook_mail_list` / `_read` / `_search` surface `inboxLink` and `isDraft`.
  - `outlook_mail_read` defaults to plain-text body (`preferText: true`).
  - `outlook_calendar_show` defaults to plain-text body (`preferText: true`).
  - New error envelope: `{ error: 'mail_quarantined', message, hint, availableAt, receivedDateTime }`.
  - 10 tools removed (`outlook_mail_move/_delete/_mark/_flag/_importance`, `outlook_calendar_create/_update/_delete/_respond`, `outlook_contacts_list`) — and the corresponding entries removed from `openclaw.plugin.json` `contracts.tools[]`.

  ## `@alavida-ai/outlook-cli` (minor)
  - 10 commands removed (`mail {move,delete,mark,flag,importance}`, `calendar {create,update,delete,respond}`, `contacts list`).
  - `mail list` / `mail read` JSON output gains `isDraft` + `inboxLink`.
  - `mail list` pretty output adds an `open: <inboxLink>` line per row.
  - `mail read` pretty output prints `Open in Outlook: <inboxLink>`.
  - `calendar show` defaults to plain-text body; new `--html` flag for raw HTML.
  - `MailQuarantinedError` rendered as `Mail blocked: …\n  Received: …\n  Available at: …`.

### Patch Changes

- Updated dependencies [f55d279]
- Updated dependencies [bfa8f53]
- Updated dependencies [2bd3ebc]
  - @alavida-ai/outlook-core@0.1.0

## 0.0.3

### Patch Changes

- 7900cc0: Rename the default token-cache directory from `~/.outlook-cli/` to `~/.outlook-plugin/` so the runtime path matches the renamed repo (`outlook-plugin`) instead of one of the three packages.
  - `@alavida-ai/outlook-cli` — `defaultCachePath()` returns `~/.outlook-plugin/tokens.json`; `auth login --help` text and `OUTLOOK_TOKEN_CACHE` docs updated
  - `@alavida-ai/outlook-plugin-openclaw` — `defaultCachePath()` in `client.ts` matches; `openclaw.plugin.json` config schema description matches; `index.ts` typebox description matches
  - `@alavida-ai/outlook-core` — `AuthLockTimeoutError.nextStep` text references the new lock path

  No backward-compat shim — nothing has been deployed to production yet.

  Operators who already ran `outlook auth login` against the legacy path should either `mv ~/.outlook-cli ~/.outlook-plugin` or re-run `outlook auth login` to write the cache at the new path.

- Updated dependencies [7900cc0]
  - @alavida-ai/outlook-core@0.0.2

## 0.0.2

### Patch Changes

- fdca89f: Rewrite the bundled skill (`skills/outlook/`) for OpenClaw tools instead of legacy CLI commands. Every example in `SKILL.md` and the five reference files (`mail.md`, `calendar.md`, `safety.md`, `auth.md`, `body-input.md`) now uses tool calls (e.g. `mail_list({ unread: true })`) instead of shell invocations (e.g. `outlook mail list -u --json`). Tool param names verified against the actual typebox schemas — fixes a few mismatches between the original CLI-era doc and the real plugin surface (`mail_read.preferText`, `mail_reply.all`, `calendar_list` taking `after/before/limit` not `days`, etc.).

  The shared `output: 'pretty' | 'json'` param is now described as **token-efficient vs more detailed** instead of human-readable vs machine-readable — both consumers are LLM agents, so the meaningful tradeoff is tokens vs detail.

  Skill metadata's `requires.bins: ["outlook"]` removed — the plugin is self-contained for tool calls, and host-side auth setup lives in `references/auth.md`, not the manifest.

  Code surface unchanged; this is a docs-only bump so updated chibote installs pick up the new bundled skill content.
