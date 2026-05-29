# @alavida-ai/outlook-plugin-openclaw

OpenClaw plugin: read and triage Microsoft Outlook mail + calendar via Microsoft Graph as the signed-in user. Delegated permissions; draft-only mail.

## Install (per OpenClaw host)

One-time `~/.npmrc` for the `@alavida-ai` scope:

```ini
@alavida-ai:registry=https://npm.pkg.github.com
//npm.pkg.github.com/:_authToken=<github-pat-with-read:packages>
```

Then:

```bash
openclaw plugins install @alavida-ai/outlook-plugin-openclaw
openclaw gateway restart
openclaw plugins inspect outlook   # confirm "Format: openclaw" + tool list
```

## First-time auth (per user, per host)

The plugin reads MSAL tokens from `~/.outlook-plugin/tokens.json`. To populate that cache on the OpenClaw host, install the matching CLI and run device-code sign-in once:

```bash
npm install -g @alavida-ai/outlook-cli
outlook auth login   # device-code → URL + code printed to stderr → sign in on any device
outlook whoami       # verify
```

After this, the OpenClaw plugin can call Graph silently as the signed-in user. Refresh tokens auto-renew; re-login is only needed on password change, admin consent revocation, or ~90 days idle.

## Plugin config (in `~/.openclaw/openclaw.json`)

```jsonc
{
  "plugins": {
    "entries": {
      "outlook": {
        "enabled": true,
        "config": {
          // All optional — defaults work for most setups.
          // "clientId":       "${AZURE_CLIENT_ID}",      // override embedded Entra app id
          // "tenantId":       "${AZURE_TENANT_ID}",      // override default 'common'
          // "tokenCachePath": "/var/lib/openclaw/outlook-tokens.json",
          // "account":        "user@tenant.onmicrosoft.com"  // for multi-account hosts
        }
      }
    }
  }
}
```

## Tenant admin consent (per client tenant, one time)

If the tenant has user-consent restrictions on (most enterprises do), an admin must consent to the app once. Send them this URL with `{tenant}` replaced by their tenant domain or id:

```
https://login.microsoftonline.com/{tenant}/adminconsent?client_id=18f9e6ff-2b0a-423e-bb35-ab9b541e604e
```

The admin signs in, reviews the requested scopes (Mail.ReadWrite, Calendars.ReadWrite, Calendars.ReadWrite.Shared, Contacts.ReadWrite, User.Read, offline_access), and clicks **Accept**. After that any user in the tenant can complete `outlook auth login` without seeing a consent prompt.

## Tools

24 native tools across four areas:

| Category | Tools |
|---|---|
| **Identity** | `whoami` |
| **Mail read** | `mail_list`, `mail_read`, `mail_search`, `mail_folders`, `mail_list_attachments`, `mail_download_attachment` |
| **Mail write (drafts only — never sends)** | `mail_draft`, `mail_reply`, `mail_forward`, `mail_add_attachment` |
| **Mail triage** | `mail_move`, `mail_delete`, `mail_mark`, `mail_flag`, `mail_importance` |
| **Calendar** | `calendar_list`, `calendar_show`, `calendar_create`, `calendar_update`, `calendar_delete`, `calendar_respond`, `calendar_availability` |
| **Contacts** | `contacts_list` (stub) |

Every tool inherits shared `output: 'pretty' \| 'json'` and `help: boolean` params. See the bundled `skills/outlook/SKILL.md` for full agent guidance.

## Source

https://github.com/alavida-ai/outlook-plugin
