---
"@alavida-ai/outlook-plugin-openclaw": patch
---

Fix browser-flow auth docs: register the callback redirect URI under **Mobile
and desktop applications** (public-client/Native type), not **Web**. The plugin
is a public client (PKCE, no secret); a Web redirect makes Entra treat it as
confidential and fail the token exchange with AADSTS7000218. Adds a note
explaining why, and that the "Allow public client flows" toggle does not apply
to the auth-code flow.
