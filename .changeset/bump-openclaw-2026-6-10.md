---
"@alavida-ai/outlook-plugin-openclaw": patch
---

Build against openclaw `2026.6.10` (latest stable; was `2026.5.12`). The newer
plugin SDK surfaced a declaration-emit portability error (TS2742) on the default
export, fixed by annotating it with the public `OpenClawPluginDefinition` type
instead of the SDK-internal inferred return type. No runtime behavior change.
The `peerDependencies` / `minGatewayVersion` floor is intentionally left at
`>=2026.3.24-beta.2` so the plugin keeps loading on older gateways.
