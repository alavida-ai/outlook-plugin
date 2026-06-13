# Auth

Each agent maintains its own Outlook identity. The plugin keeps a separate
token cache per agent (at `<agentDir>/outlook-tokens.json`), so two agents
on the same OpenClaw gateway can sign in as two different Microsoft users
without crossing wires.

## Two sign-in flows

`outlook_auth_login` supports two flows. **The plugin picks automatically
based on config** — the agent surface is the same either way:

| Flow | When it's used | What the agent surfaces |
| --- | --- | --- |
| **Browser (Authorization Code + PKCE)** | `oauthRedirectUri` is set in plugin config | A single sign-in `authUrl` |
| **Device code** | `oauthRedirectUri` is unset | A `verificationUrl` + 6-character `userCode` |

The **browser flow is the primary path** and is required for tenants whose
Conditional Access blocks device-code flow (Microsoft's recommended
"Block Device Code Flow" baseline). The **device-code flow is the fallback**
for clients without that restriction and for local development.

Both flows return **immediately** with `status: "pending"` — neither blocks
for sign-in — and both finish the same way: the human signs in, tokens land
in this agent's cache, and you confirm with `outlook_auth_status`.

## Authenticating an agent (from inside the session)

Call `outlook_auth_login` from the agent.

**Browser flow** (when `oauthRedirectUri` is configured):

```jsonc
{
  "status": "pending",
  "flow": "browser",
  "authUrl": "https://login.microsoftonline.com/<tenant>/oauth2/v2.0/authorize?...",
  "expiresAt": "2026-06-01T12:10:00Z",
  "agentId": "alfred",
  "cachePath": "/.../agents/alfred/agent/outlook-tokens.json",
  "hint": "Open the URL in a browser and sign in. Then call outlook_auth_status to confirm."
}
```

Surface `authUrl` to the human. They open it in any browser, sign in
(subject to the tenant's MFA / Conditional Access), and Microsoft redirects
their browser to the plugin's `/outlook/auth-callback` route, which redeems
the code and writes tokens to **this agent's** cache. The link is
single-use and expires after 10 minutes.

**Device-code flow** (when `oauthRedirectUri` is unset):

```jsonc
{
  "status": "pending",
  "flow": "device",
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
Polling Microsoft happens in the background.

In **both** flows, once the human confirms sign-in is done, call
`outlook_auth_status` to verify the token landed.

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
them to try again — browser links expire after 10 min, device codes after
~15 min) or Microsoft rejected the sign-in. The pending flow is single-use,
so a fresh `outlook_auth_login` is needed for each attempt. Surface the
failure and ask them to retry `outlook_auth_login`.

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

## Enabling the browser flow (operator config)

Set `oauthRedirectUri` in plugin config to switch `outlook_auth_login` to
the browser (Authorization Code + PKCE) flow. The value is the **public
HTTPS callback URL** Microsoft will redirect to, and it must **exactly
match** a redirect URI registered in the Entra app (no wildcards).

```json
{
  "plugins": {
    "outlook": {
      "oauthRedirectUri": "https://<gateway>.<tailnet>.ts.net/outlook/auth-callback"
    }
  }
}
```

The plugin serves the callback at `/outlook/auth-callback` (plugin-scoped
auth, exact-match). To expose **only** that path publicly — keeping the rest
of the gateway Tailnet-only — run on the gateway box:

```bash
tailscale funnel --bg --set-path /outlook/auth-callback \
  http://127.0.0.1:18789/outlook/auth-callback
tailscale funnel status   # prints the public *.ts.net URL
```

Tailscale auto-issues and auto-renews the TLS cert for the `*.ts.net` URL;
end users do not need to be on the Tailnet. Then register that exact URL as
a redirect URI in the Entra app (Authentication → Platform configurations →
Add a platform → **Mobile and desktop applications** → enter it under
**Custom redirect URIs**).

> ⚠️ Register the callback under **Mobile and desktop applications**, **not
> Web**. The plugin authenticates as a *public client* (PKCE, no client
> secret). Entra decides public vs confidential by the redirect URI's
> platform type: a **Web** redirect makes Entra treat the app as confidential
> and demand a `client_secret`, so the token exchange fails with
> **AADSTS7000218** ("request body must contain 'client_assertion' or
> 'client_secret'"). A **Mobile and desktop applications** (Native) redirect
> classifies the app as public — no secret needed. The "Allow public client
> flows" toggle does **not** fix this: it is only consulted when the token
> request carries no redirect URI (device-code / ROPC), so it never applies
> to the auth-code flow. If you later move to a confidential client (with a
> secret), the Web platform becomes the correct choice instead.

Leave `oauthRedirectUri` unset to keep the device-code flow.

### Browser-flow security model

- **PKCE** — the code verifier never leaves the gateway; only its SHA-256
  challenge is sent to Microsoft. The verifier is never returned to the
  agent or written to any log.
- **state** — a random CSRF token tied server-side to the initiating agent.
  Single-use and 10-minute TTL; a replayed or expired callback is refused.
- **nonce** — bound into the issued ID token and verified on receipt.
- **Path isolation** — only `/outlook/auth-callback` is funnelled; every
  other gateway path stays Tailnet-only.

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
