---
"@alavida-ai/outlook-plugin-openclaw": minor
"@alavida-ai/outlook-core": minor
"@alavida-ai/outlook-cli": minor
---

Auth cleanup: remove the device-code flow and drop CLI multi-account support.

- **core**: remove `loginDeviceCode` / `loginDeviceCodeInBackground` (and their
  types); the `LoginResult` type moves to its own module and is still exported.
- **openclaw**: `outlook_auth_login` is now browser-only (Authorization Code +
  PKCE) and **requires** `oauthRedirectUri` — it returns a clear error if unset.
  Device-code is gone (it's blocked by the Conditional Access baselines we
  target, and the localhost interactive flow can't run on a headless gateway).
- **cli**: single-account. `--account` / `OUTLOOK_ACCOUNT` are removed and
  `outlook auth login` now clears any cached account before signing in, so the
  CLI cache always holds exactly one identity.
