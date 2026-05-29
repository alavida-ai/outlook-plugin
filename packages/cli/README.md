# @alavida-ai/outlook-cli

Terminal CLI: read mail, draft messages, manage calendar via Microsoft Graph.

Stdout = data. Stderr = human messages. Exit 0 success, 1 user/auth error, 2 unexpected.

## Install

```bash
npm install -g @alavida-ai/outlook-cli
```

Requires the `@alavida-ai` scope in your `~/.npmrc`:

```ini
@alavida-ai:registry=https://npm.pkg.github.com
//npm.pkg.github.com/:_authToken=<github-pat-with-read:packages>
```

## First-time auth

```bash
outlook auth login   # device-code → URL + code on stderr → sign in on any device
outlook whoami       # verify
```

Tokens cache to `~/.outlook-plugin/tokens.json` (0600). Refresh tokens auto-renew; re-login only needed on password change, admin consent revocation, or ~90 days idle.

## Commands

| Group | Commands |
|---|---|
| **Auth** | `auth login`, `auth logout`, `auth status` |
| **Identity** | `whoami` |
| **Mail (read)** | `mail list`, `mail read <id>`, `mail search "<query>"`, `mail folders`, `mail attachments <id>`, `mail download-attachment <id> <att-id>` |
| **Mail (write — drafts only)** | `mail draft`, `mail reply <id>`, `mail forward <id>`, `mail add-attachment <id>` |
| **Mail (triage)** | `mail move <id>`, `mail delete <id>`, `mail mark <id>`, `mail flag <id>`, `mail importance <id>` |
| **Calendar** | `calendar list`, `calendar show <id>`, `calendar create`, `calendar update <id>`, `calendar delete <id>`, `calendar respond <id>`, `calendar availability` |
| **Contacts** | `contacts list` (stub) |

Run `outlook <command> --help` for per-command flags. Every command supports `--json`.

## Global flags

- `--account <UPN>` — pick a specific cached account (or set `OUTLOOK_ACCOUNT`)
- `--json` — emit JSON to stdout instead of a human summary

## Environment

- `AZURE_CLIENT_ID` — override the embedded Entra app id
- `AZURE_TENANT_ID` — override the default tenant (`common`)
- `OUTLOOK_ACCOUNT` — default UPN to use when multiple accounts cached
- `OUTLOOK_TOKEN_CACHE` — override `~/.outlook-plugin/tokens.json`

## Source

https://github.com/alavida-ai/outlook-plugin
