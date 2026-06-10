# Auth

Each agent maintains its own Outlook identity. The plugin keeps a separate
token cache per agent (at `<agentDir>/outlook-tokens.json`), so two agents
on the same OpenClaw gateway can sign in as two different Microsoft users
without crossing wires.

## Authenticating an agent (from inside the session)

Call `outlook_auth_login` from the agent. The tool returns **immediately**
with a verification URL and a 6-character user code — it does **not** block
for sign-in:

```jsonc
{
  "status": "pending",
  "verificationUrl": "https://microsoft.com/devicelogin",
  "userCode": "ABCD1234",
  "expiresAt": "2026-06-01T12:15:00Z",
  "agentId": "alfred",
  "cachePath": "/.../agents/alfred/agent/outlook-tokens.json",
  "hint": "Open the URL on any device, enter the code, sign in. Then call outlook_auth_status to confirm."
}
```

Surface `verificationUrl` and `userCode` to the human. The human signs in
on any device (phone is fine), enters the code, and authorises sign-in.
Polling Microsoft happens in the background; once the human confirms
sign-in is done, call `outlook_auth_status` to verify the token landed.

```jsonc
{
  "status": "authenticated",
  "upn": "alice@example.com",
  "displayName": "Alice Smith",
  "agentId": "alfred",
  "cachePath": "/.../agents/alfred/agent/outlook-tokens.json"
}
```

If `auth_status` still returns `not_authenticated` after the human says
sign-in is done, either the human hasn't actually completed the flow (ask
them to try again — codes expire after ~15 min) or Microsoft rejected the
sign-in. Surface the failure and ask them to retry `outlook_auth_login`.

## Logging out

```text
outlook_auth_logout
```

Clears this agent's token cache. The next tool call returns
`auth_cache_missing` until `outlook_auth_login` runs again. The cache for
any other agent on the same gateway is untouched.

## When tools fail with an auth error

Every Outlook tool returns a structured error envelope if auth is missing
or stale:

```json
{ "__toolError": { "error": "<code>", "message": "...", "hint": "..." } }
```

The agent's job is to **surface `hint` to the human** and either retry
after they've re-authed or stop. Don't loop on auth errors.

### Auth error taxonomy

| `error` code | Meaning | Recovery |
| --- | --- | --- |
| `auth_cache_missing` | No cached account for this agent | Call `outlook_auth_login` |
| `auth_cache_corrupt` | Token cache unreadable | Call `outlook_auth_logout` then `outlook_auth_login` |
| `auth_refresh_failed` | Silent refresh rejected (password change, revoked consent, conditional-access re-eval) | Call `outlook_auth_login` |
| `auth_interaction_required` | Microsoft requires interactive sign-in (MFA prompt, etc.) | Call `outlook_auth_login` |
| `auth_ambiguous_account` | Multiple accounts cached, none selected | Pin `account: '<upn>'` in plugin config (see below) |
| `auth_lock_timeout` | Another process holds the token-cache refresh lock | Wait a few seconds and retry; if persistent, call `outlook_auth_logout` and `outlook_auth_login` |

`auth_ambiguous_account` includes the cached UPNs in the envelope's
`accounts` field — surface that list so the operator knows which accounts
are available.

## Account selection (operator config, not per-call)

The agent does **not** pick the account at call time. Account pinning is
plugin configuration the operator sets once in
`~/.openclaw/openclaw.json`:

```json
{
  "plugins": {
    "outlook": {
      "account": "user@example.com"
    }
  }
}
```

When only one account is cached, this field is optional. When multiple
accounts are cached and `account` is unset, tools throw
`auth_ambiguous_account`.

## Multi-agent isolation

On a multi-agent OpenClaw gateway (`alfred`, `baerbel`, …), each agent's
token cache lives at `<agentDir>/outlook-tokens.json` (e.g.
`~/.openclaw/agents/alfred/agent/outlook-tokens.json`). The cache file is
`0600`. Tools running inside `alfred`'s session never see `baerbel`'s
tokens.

`outlook_auth_login` from inside `alfred` writes to `alfred`'s cache only.
`outlook_auth_logout` from inside `alfred` deletes `alfred`'s cache only.

If the operator overrides `tokenCachePath` in plugin config, all agents
share that single file — only useful for single-agent hosts.

## Re-auth triggers

Tokens silently refresh forever, except in these cases — all of which
surface as `auth_refresh_failed` or `auth_interaction_required`, and
require the agent to call `outlook_auth_login` again:

- **Password change** → all refresh tokens invalidated server-side
- **Admin consent revocation** → same
- **90-day continuous idle** → refresh token expires
- **Conditional Access re-evaluation** → may force re-login
- **User signs out at https://myaccount.microsoft.com** → tokens invalidated

These are uncommon. Default assumption: once an agent has authed, it stays
authed.

## CLI authentication (host operator only)

The host operator can also run `outlook auth login` from the terminal. The
CLI uses a **separate** cache at `~/.outlook-plugin/tokens.json`, isolated
from any agent's cache. Running `outlook auth login` does **not**
authenticate any OpenClaw agent — the agent must call its own
`outlook_auth_login` tool.

```bash
outlook auth login                 # CLI auth (host-side, separate from agents)
outlook auth status                # CLI status
outlook auth logout                # CLI logout
outlook whoami                     # CLI: who am I (CLI cache only)
```

This is a deliberate posture: the CLI is for the human host operator
(debugging, status checks); the plugin tools are for agents.

## Token storage

| Platform | Backend | Notes |
| --- | --- | --- |
| macOS | File (`0600`) | OS keychain support is future work |
| Linux | File (`0600`) | OS keychain support is future work |
| Windows | File (`0600`) | OS keychain support is future work |

The cache file is `0600`, the parent directory is `0700`. Neither the CLI
nor the plugin ever asks for or stores Microsoft passwords — only access +
refresh tokens issued by Microsoft via OAuth.
