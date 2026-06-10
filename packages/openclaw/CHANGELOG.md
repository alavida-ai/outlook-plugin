# @alavida-ai/outlook-plugin-openclaw

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
