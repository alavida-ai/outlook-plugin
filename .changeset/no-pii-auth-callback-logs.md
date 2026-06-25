---
"@alavida-ai/outlook-plugin-openclaw": patch
---

Stop logging the signed-in user's email (`upn`) in the auth-callback success
line. The handler now records `sign-in complete for agent=<id>` only, so no
personal data lands in gateway server logs. The token exchange still returns
the `upn` internally; it is simply no longer written to stderr.
