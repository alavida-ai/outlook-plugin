# @alavida-ai/outlook-plugin-openclaw

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
