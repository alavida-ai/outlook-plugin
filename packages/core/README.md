# @alavida-ai/outlook-core

Pure TypeScript library: MSAL device-code auth + Microsoft Graph client wiring for Alavida's Outlook stack.

Consumed by:

- [`@alavida-ai/outlook-cli`](https://github.com/alavida-ai/outlook-cli/tree/typescript-rewrite/packages/cli) — the `outlook` terminal binary
- [`@alavida-ai/outlook-plugin-openclaw`](https://github.com/alavida-ai/outlook-cli/tree/typescript-rewrite/packages/openclaw) — the OpenClaw plugin

Both share the same token cache (`~/.outlook-cli/tokens.json` by default) — sign in via either, the other picks it up.

## Public surface

- **`OutlookClient`** + resources (currently `.me`; mail/calendar/contacts in subsequent slices)
- **Auth functions:** `loginDeviceCode`, `getAccessToken`, `logout`, `status`, `resolveAccount`
- **Token cache:** `TokenCache` interface + `FileTokenCache` implementation (atomic writes, cross-process `O_EXCL` lock, integrity check, defined error taxonomy)
- **MSAL factory:** `buildMsalApp`, `EMBEDDED_CLIENT_ID`, `OUTLOOK_SCOPES`
- **Graph wiring:** `makeGraphClient`, `MsalAuthenticationProvider`, `liftGraphError`
- **Error taxonomy:** `AuthError` family + `CoreError` / `NotFoundError` / `ThrottledError` / `ServerError` / `NetworkError`

## Source

https://github.com/alavida-ai/outlook-cli
