---
'@alavida-ai/outlook-core': minor
'@alavida-ai/outlook-plugin-openclaw': minor
'@alavida-ai/outlook-cli': patch
---

Per-agent token cache + plugin-side auth tools.

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
