# Auth (CLI)

The CLI maintains its own token cache at `~/.outlook-plugin/tokens.json` (mode `0600`). This is **separate** from any openclaw plugin agent's cache. Running `outlook auth login` from the terminal does **not** authenticate any agent, and vice versa.

## Sign in

```bash
outlook auth login
```

This uses the **Authorization Code + PKCE** flow: your default browser opens to the Microsoft sign-in page and MSAL listens on a `http://localhost` loopback for the redirect. Nothing needs to be exposed publicly and no redirect URI has to be registered — Microsoft accepts localhost for public clients. This flow is compatible with Conditional Access policies that block device-code sign-in.

If the browser can't be opened automatically (e.g. you're over SSH), the URL is printed to stderr so you can open it yourself:

```
Opening your browser to sign in. If it doesn't open, visit:
https://login.microsoftonline.com/<tenant>/oauth2/v2.0/authorize?...
```

Sign in with the Microsoft account (work/school or personal) and approve the consent screen. The consent screen will show the requested scopes:

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

Both should print your UPN, display name, and tenant info. If `auth status` says "not authenticated" after you completed sign-in in the browser, the sign-in may have timed out or been cancelled — re-run `outlook auth login`.

## Log out

```bash
outlook auth logout
```

Clears the cached tokens. Next command that hits Graph will fail until you log back in.

## One account at a time

The CLI is single-account: `outlook auth login` **replaces** whoever was signed in (it clears the cache before signing in), so you never juggle accounts or pass a `--account` flag. To switch identities, just run `outlook auth login` again and sign in as the other user.

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
| `auth_ambiguous_account` (a pre-existing cache holds >1 account) | `outlook auth login` — it clears the cache and signs in as one account |
| `auth_lock_timeout` (another process holds the refresh lock) | Wait a few seconds and retry; if persistent, `outlook auth logout && outlook auth login` |

## Where the cache lives

| Path | Notes |
| --- | --- |
| `~/.outlook-plugin/tokens.json` | Default. Mode `0600`; parent dir `0700`. |
| `$OUTLOOK_TOKEN_CACHE` (env var) | Override for testing / VPS-specific layouts. |

The cache contains MSAL-format access tokens, refresh tokens, ID tokens, and account metadata. No passwords are stored — only OAuth tokens issued by Microsoft.
