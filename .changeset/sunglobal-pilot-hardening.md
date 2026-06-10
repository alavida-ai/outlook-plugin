---
'@alavida-ai/outlook-core': minor
'@alavida-ai/outlook-plugin-openclaw': minor
'@alavida-ai/outlook-cli': minor
---

Sun Global pilot hardening: scope downgrade, 30-min mail safety window, cloud.microsoft inbox URLs, plain-text body defaults, and `outlook_` prefix on all openclaw tool names.

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
