# Auth

The agent never authenticates the user itself. Auth is a **one-time host-side setup** the operator performs on the OpenClaw host before the agent ever runs.

## One-time host-side login (operator action)

```bash
outlook auth login
```

What happens, step by step:

1. CLI hits Microsoft's `/devicecode` endpoint, gets a short URL + 6-character user code (~200 ms)
2. **Prints URL + code to stderr**, then blocks polling Microsoft every ~5 seconds for up to 15 minutes
3. The operator (or end-user) opens the URL on any device (phone is fine), enters the code, signs in with normal Microsoft credentials (MFA if enforced)
4. CLI unblocks, caches tokens to OS keychain (or 0600 file fallback), exits 0

After that, every OpenClaw tool call silently uses the cached tokens.

## Agent's role: surface, don't retry

Tools never spawn `outlook auth login`. When an auth-time failure happens, the agent gets a structured error envelope:

```json
{ "__toolError": { "error": "<code>", "message": "...", "hint": "..." } }
```

The agent's job is to **surface `hint` to the human** and stop. Do not retry; do not attempt a workaround. The operator must run `outlook auth login` on the host to recover.

### Auth error taxonomy

| `error` code | Meaning | Operator action |
| --- | --- | --- |
| `auth_cache_missing` | No cached account on this host | Run `outlook auth login` |
| `auth_cache_corrupt` | Token cache unreadable | Run `outlook auth logout` then `outlook auth login` |
| `auth_refresh_failed` | Silent refresh rejected (password change, revoked consent, conditional-access re-eval) | Run `outlook auth login` |
| `auth_interaction_required` | Microsoft requires interactive sign-in (MFA prompt, etc.) | Run `outlook auth login` |
| `auth_ambiguous_account` | Multiple cached accounts, none selected | Pin `account: '<upn>'` in plugin config (see below) |
| `auth_lock_timeout` | Another process holds the token-cache refresh lock | Wait a few seconds and retry; if persistent, the operator can delete `~/.outlook-plugin/tokens.lock` |

`auth_ambiguous_account` includes the cached UPNs in the envelope's `accounts` field — surface that list to the user so they (or the operator) know which accounts are available.

## Account selection (operator config, not per-call)

The agent does **not** pick the account at call time. Account pinning is plugin configuration the operator sets once in `~/.openclaw/openclaw.json`:

```json
{
  "plugins": {
    "outlook": {
      "account": "user@example.com"
    }
  }
}
```

When only one account is cached, this field is optional. When multiple accounts are cached and `account` is unset, tools throw `auth_ambiguous_account`.

## Detecting auth state (operator use)

```bash
outlook auth status                # exits 0 if cached account exists, 1 if not
outlook whoami                     # exits 1 with a friendly stderr message if not authed
```

These are CLI commands the operator runs on the host — the agent doesn't call them. The agent simply waits for any tool to throw `auth_cache_missing` (or another `auth_*` code) and surfaces the hint.

## Re-auth triggers

Tokens silently refresh forever, except in these cases — all of which surface as `auth_refresh_failed` or `auth_interaction_required` for the agent, and require an operator-side `outlook auth login`:

- **Password change** → all refresh tokens invalidated server-side
- **Admin consent revocation** → same
- **90-day continuous idle** → refresh token expires
- **Conditional Access re-evaluation** → may force re-login
- **User signs out at https://myaccount.microsoft.com** → tokens invalidated

These are uncommon. Default assumption: once the operator authed, the tools stay authed.

## Logout (operator action)

```bash
outlook auth logout
```

Clears tokens from OS keychain + file fallback. The next tool call returns `auth_cache_missing` until the operator runs `outlook auth login` again.

## Token storage

| Platform | Backend | Notes |
| --- | --- | --- |
| macOS | Keychain | Native, no setup |
| Linux desktop | Secret Service (libsecret) | Requires GNOME Keyring or KDE Wallet |
| Windows | Credential Manager | Native, no setup |
| Headless Linux (e.g. VPS) | File at `~/.outlook-plugin/tokens.json` | 0600 perms, parent dir 0700 |

The CLI tries the OS keychain first; if none is available, it falls back to the encrypted-at-rest file. Neither the CLI nor the plugin ever asks for or stores Microsoft passwords — only access + refresh tokens issued by Microsoft via OAuth.

## Re-triggering auth from an unusual context

If the agent runs in an environment where it must hand off auth setup to a remote user (e.g. provisioning a new user mid-conversation), spawn the CLI as a subprocess on the host and forward the first line of stderr — that carries the URL + code:

```python
import subprocess
proc = subprocess.Popen(
    ['outlook', 'auth', 'login'],
    stderr=subprocess.PIPE, stdout=subprocess.DEVNULL, text=True,
)
url_line = proc.stderr.readline()
# url_line: "To sign in, use a web browser to open ... and enter the code ABCD1234 ..."
send_to_user_via_preferred_channel(url_line)
returncode = proc.wait()  # blocks up to 15 min
```

This is the **exception path**, not the default. The default is: operator runs `outlook auth login` once at host setup time, agent never thinks about auth again.
