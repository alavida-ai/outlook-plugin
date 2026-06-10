---
'@alavida-ai/outlook-plugin-openclaw': patch
---

Fix every tool showing as `(anonymous)` in `openclaw plugins inspect outlook --runtime`.

The plugin registers each tool as a **factory** (so the per-agent token cache can be baked into the config at agent-setup time, not module load time). When you register a factory, openclaw cannot introspect the tool's name without invoking the factory — and the inspector deliberately doesn't invoke factories at inspect time. The SDK provides a separate `opts.name` argument on `api.registerTool` for exactly this case; we were not passing it.

`registerTool` in `packages/openclaw/src/register.ts` now passes `{ name: descriptor.name }` as the second argument. After updating, `openclaw plugins inspect outlook --runtime` will display the real tool names (`outlook_mail_list`, `outlook_auth_login`, etc.) instead of 17 lines of `(anonymous)`.

This is a cosmetic fix — tools were callable from agents before this change. Worth shipping because the inspector display is the primary way operators sanity-check that a plugin loaded correctly, and because the `optional`/manifest-alignment flow depends on the same name hint.
