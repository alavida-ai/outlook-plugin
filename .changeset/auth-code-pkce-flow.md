---
"@alavida-ai/outlook-plugin-openclaw": minor
"@alavida-ai/outlook-core": minor
"@alavida-ai/outlook-cli": minor
---

Add the Authorization Code + PKCE sign-in flow alongside device-code.

- **core**: new `buildAuthCodeUrl` / `exchangeAuthCode` helpers (PKCE verifier,
  CSRF `state`, ID-token `nonce`, nonce verification on exchange) and a
  `loginInteractive` helper for the localhost-loopback interactive flow.
- **openclaw**: `outlook_auth_login` returns a browser sign-in URL when the new
  `oauthRedirectUri` plugin config is set, and a new `/outlook/auth-callback`
  HTTP route (plugin-scoped, exact-match) redeems the code into the initiating
  agent's token cache. Pending flows are single-use with a 10-minute TTL. When
  `oauthRedirectUri` is unset, the existing device-code flow is unchanged.
- **cli**: `outlook auth login` now uses the interactive browser flow
  (`acquireTokenInteractive`) instead of device-code, so it works on tenants
  whose Conditional Access blocks device-code sign-in.
