---
'@alavida-ai/outlook-core': patch
'@alavida-ai/outlook-cli': patch
'@alavida-ai/outlook-plugin-openclaw': patch
---

Rename the default token-cache directory from `~/.outlook-cli/` to `~/.outlook-plugin/` so the runtime path matches the renamed repo (`outlook-plugin`) instead of one of the three packages.

- `@alavida-ai/outlook-cli` — `defaultCachePath()` returns `~/.outlook-plugin/tokens.json`; `auth login --help` text and `OUTLOOK_TOKEN_CACHE` docs updated
- `@alavida-ai/outlook-plugin-openclaw` — `defaultCachePath()` in `client.ts` matches; `openclaw.plugin.json` config schema description matches; `index.ts` typebox description matches
- `@alavida-ai/outlook-core` — `AuthLockTimeoutError.nextStep` text references the new lock path

No backward-compat shim — nothing has been deployed to production yet.

Operators who already ran `outlook auth login` against the legacy path should either `mv ~/.outlook-cli ~/.outlook-plugin` or re-run `outlook auth login` to write the cache at the new path.
