# Auth (CLI)

The CLI maintains its own token cache at `~/.outlook-plugin/tokens.json` (mode `0600`). This is **separate** from any openclaw plugin agent's cache. Running `outlook auth login` from the terminal does **not** authenticate any agent, and vice versa.

## Sign in

```bash
outlook auth login
```

Prints something like:

```
To sign in, use a web browser to open https://microsoft.com/devicelogin and enter the code ABCD1234 to authenticate.
```

On any device, open the URL, paste the code, sign in with the Microsoft account (work/school or personal), approve the consent screen. The consent screen will show the requested scopes:

- Read your mail (Mail.ReadWrite — draft creation needs the W half; no send)
- Read your calendars (Calendars.Read)
- Read your calendars in shared folders (Calendars.Read.Shared)
- Sign you in and read your profile (User.Read)
- Maintain access to data you have given it access to (offline_access)

If consent is required for the tenant (work/school account), the tenant admin may need to approve the app — `outlook auth login` will surface the admin-consent URL if so.

## Verify

```bash
outlook auth status
outlook whoami
```

Both should print your UPN, display name, and tenant info. If `auth status` says "not authenticated" after you completed sign-in in the browser, the code may have expired (~15 min) — re-run `outlook auth login`.

## Log out

```bash
outlook auth logout
```

Clears the cached tokens. Next command that hits Graph will fail until you log back in.

## Switching accounts

The CLI supports multiple cached accounts. Pin a specific account per-call with `--account`:

```bash
outlook --account alice@example.com mail list
outlook --account bob@other.com mail list
```

Or set a default for the shell session:

```bash
export OUTLOOK_ACCOUNT=alice@example.com
outlook mail list   # uses alice
```

If there are multiple cached accounts and neither flag nor env var is set, every command will fail with `auth_ambiguous_account` listing the available UPNs.

## Re-auth triggers (uncommon)

Tokens silently refresh forever, except:

- **Password change** → all refresh tokens invalidated server-side
- **Admin consent revocation** → same
- **90-day continuous idle** → refresh token expires
- **Conditional Access re-evaluation** → may force re-login
- **User signed out at `myaccount.microsoft.com`** → tokens invalidated
- **Scope change in the CLI** → existing token's grant doesn't match the new scope set; need fresh consent

When any of these fires, the next command will emit a clear "next step" hint pointing you at `outlook auth login`.

## Common issues

| Problem | Fix |
| --- | --- |
| `auth_cache_missing` on a fresh install | `outlook auth login` |
| `auth_cache_corrupt` after upgrade | `outlook auth logout && outlook auth login` |
| `auth_refresh_failed` after password change or scope change | `outlook auth login` |
| `auth_ambiguous_account` | Pin `--account UPN` or `export OUTLOOK_ACCOUNT=...` |
| `auth_lock_timeout` (another process holds the refresh lock) | Wait a few seconds and retry; if persistent, `outlook auth logout && outlook auth login` |

## Where the cache lives

| Path | Notes |
| --- | --- |
| `~/.outlook-plugin/tokens.json` | Default. Mode `0600`; parent dir `0700`. |
| `$OUTLOOK_TOKEN_CACHE` (env var) | Override for testing / VPS-specific layouts. |

The cache contains MSAL-format access tokens, refresh tokens, ID tokens, and account metadata. No passwords are stored — only OAuth tokens issued by Microsoft.
