---
'@alavida-ai/outlook-core': minor
'@alavida-ai/outlook-cli': minor
'@alavida-ai/outlook-plugin-openclaw': minor
---

Initial TypeScript rewrite — full v1 surface.

Replaces the Python `outlook-cli` with a TypeScript pnpm monorepo across three packages, at parity with the Python feature set plus a real OpenClaw plugin and stronger auth robustness.

**`@alavida-ai/outlook-core`** — pure library

- MSAL device-code auth with multi-account safety (never silently picks `accounts[0]`)
- File-based token cache (`~/.outlook-cli/tokens.json`) with atomic writes (tmpfile+fsync+rename), cross-process `O_EXCL` refresh lock, integrity check on read, defined error taxonomy
- Microsoft Graph wiring via `@microsoft/microsoft-graph-client`
- Resources: `me`, `mail` (list/read/search/folders/attachments/draft/reply/forward/move/delete/mark/flag/importance), `calendar` (list/show/create/update/delete/respond/availability), `contacts` (stub)
- Error taxonomy lifts `GraphError` into typed `AuthError`/`NotFoundError`/`ThrottledError`/`ServerError`/`NetworkError`

**`@alavida-ai/outlook-cli`** — terminal CLI

- `outlook auth login` / `logout` / `status`
- `outlook whoami`
- `outlook mail list / read / search / folders / attachments / download-attachment / draft / reply / forward / add-attachment / move / delete / mark / flag / importance`
- `outlook calendar list / show / create / update / delete / respond / availability`
- `outlook contacts list` (stub)
- Stdout = data, stderr = humans; `--json` everywhere; `--account UPN` / `OUTLOOK_ACCOUNT` for multi-account hosts; escape-decoding for `\n`/`\r`/`\t`/`\\` in body/comment args
- No CLI-framework dep — Node's `util.parseArgs` + lazy command imports

**`@alavida-ai/outlook-plugin-openclaw`** — OpenClaw plugin

- 24 native tools registered: `whoami` + mail / calendar / contacts surface as above
- Reuses the same token cache as the CLI — sign in once on either, the other picks it up
- Plugin config: `clientId`, `tenantId`, `tokenCachePath`, `account` (all optional)
- Bundled skill at `skills/outlook/` with full agent guidance

See `docs/superpowers/specs/2026-05-28-outlook-typescript-rewrite-design.md` for the full design rationale.
