# outlook-plugin

Alavida's TypeScript monorepo for Microsoft Outlook — mail, calendar, contacts via Microsoft Graph, designed for AI agents.

Three packages:

| Package | Purpose |
|---|---|
| `@alavida-ai/outlook-core` | Microsoft Graph wrapper + MSAL device-code auth. Pure library — no stdout/stderr, no `process.exit`. |
| `@alavida-ai/outlook-cli` | The `outlook` binary. Native `node:util` `parseArgs`, no CLI-framework dep. Used for local testing and one-time auth setup on OpenClaw hosts. |
| `@alavida-ai/outlook-plugin-openclaw` | OpenClaw plugin. One `ToolDescriptor` per CLI subcommand. Bundles the agent-facing skill. |

Draft-only mail (no send), delegated permissions only, OS-keychain-free file token cache with atomic writes + cross-process refresh lock + integrity checks. See `docs/superpowers/specs/2026-05-28-outlook-typescript-rewrite-design.md` for the full design.

## Status

TypeScript port of the original Python `outlook-cli`. Foundation + 23 tools shipped (whoami, mail-read/write/triage, calendar, contacts stub). At parity with the Python CLI's surface. See [commit history](https://github.com/alavida-ai/outlook-plugin/commits/main) for the slice-by-slice rollout. Agent-facing reference at [packages/openclaw/skills/outlook/SKILL.md](packages/openclaw/skills/outlook/SKILL.md).

## Architecture

```
outlook-plugin/
├── package.json                              # workspace root, pinned pnpm@11.1.2
├── pnpm-workspace.yaml                       # supply-chain hardening
├── scripts/check-pnpm.mjs                    # preinstall guard
├── tsconfig.base.json, tsconfig.json
├── eslint.config.js, .prettierrc.json, vitest.config.ts
├── docs/superpowers/                         # spec + plan files
└── packages/
    ├── core/
    │   └── src/
    │       ├── auth/                         # MSAL, FileTokenCache, multi-account, error taxonomy
    │       ├── graph/                        # MSAL → Graph auth provider, error lift
    │       ├── resources/                    # me, mail, calendar
    │       ├── client.ts                     # OutlookClient facade
    │       └── index.ts                      # public barrel
    ├── cli/
    │   └── src/
    │       ├── commands/                     # one file per `outlook <verb>` subcommand
    │       ├── client.ts, output.ts, escapes.ts
    │       └── index.ts                      # lazy-import dispatcher
    └── openclaw/
        ├── openclaw.plugin.json              # plugin manifest
        ├── skills/outlook/                   # agent-facing skill (bundled)
        └── src/
            ├── tools/                        # one file per OpenClaw tool
            ├── register.ts                   # shared output/help injection
            ├── pretty.ts                     # shape-detected renderers
            ├── client.ts                     # memoised getClient
            └── index.ts                      # plugin entry
```

## Permissions (delegated, no app perms)

- `Mail.ReadWrite` — read + create/update drafts. **No `Mail.Send` scope.**
- `Calendars.ReadWrite`
- `Calendars.ReadWrite.Shared`
- `Contacts.ReadWrite`
- `User.Read`
- `offline_access` (refresh tokens — added automatically by MSAL)

Same scope set as the Python implementation; users migrating from Python don't re-consent.

## How auth works

Embedded multi-tenant Entra app — same client id as the Python implementation (`18f9e6ff-2b0a-423e-bb35-ab9b541e604e`), `common` authority. End users sign in once via device-code, refresh tokens renew silently from then on.

**Token cache** lives at `~/.outlook-cli/tokens.json` (0600), atomic writes via tmpfile + `fsync` + `rename`, cross-process refresh lock via `O_EXCL` on `tokens.json.lock`. Six typed `AuthError` variants surface via every CLI command and OpenClaw tool with a `nextStep` field pointing the user at the right command.

Override the embedded app id via env vars (e.g. a client running their own Entra app):

```bash
export AZURE_CLIENT_ID=<their-app-id>
export AZURE_TENANT_ID=<their-tenant-id>
```

Per spec §4 ("Auth strategy"). Full robustness commitments — atomic writes, cross-process lock, integrity checks, multi-account handling, defined error taxonomy — are in §4.2 of the design doc.

## Setup

Prerequisites: Node ≥ 20, pnpm 11 (via `corepack enable`).

```bash
corepack enable
pnpm install
pnpm build
```

The build produces:
- `packages/cli/dist/index.js` — the `outlook` binary (executable bit set)
- `packages/openclaw/dist/index.js` — the plugin entry
- `packages/core/dist/` — TypeScript declarations + sources

## CLI usage

```bash
# auth (one-time per user, refresh tokens auto-renew from then on)
node packages/cli/dist/index.js auth login
node packages/cli/dist/index.js auth status
node packages/cli/dist/index.js auth logout

# mail
node packages/cli/dist/index.js mail list -u                        # unread inbox
node packages/cli/dist/index.js mail list --from boss@co.com --json
node packages/cli/dist/index.js mail draft --to x@y.com --subject "..." --body "..."
node packages/cli/dist/index.js mail search "subject:invoice" --json

# calendar
node packages/cli/dist/index.js calendar list -d 7
node packages/cli/dist/index.js calendar availability --emails a@b.com --emails c@d.com -d 5
```

After `pnpm link --global packages/cli` (or once we publish to GitHub Packages), `outlook` is on `$PATH` and the `node packages/cli/dist/index.js` prefix goes away.

All data commands support `--json`. Stdout = data, stderr = human messages — pipe stdout to `jq` safely.

Multi-account hosts: pass `--account <upn>` or set `OUTLOOK_ACCOUNT=<upn>`. The CLI never silently picks among cached accounts.

## OpenClaw plugin

The plugin runs every CLI subcommand as an OpenClaw `ToolDescriptor` — same `core` code path, no shelling out. Twenty-three tools registered (one per CLI subcommand minus the `auth` triplet which stays CLI-only).

Distribution lives in `~/.agentkb/alavida/wiki/openclaw-plugin-distribution.md`. Once a GitHub Actions workflow publishes the plugin to GitHub Packages:

```bash
# on the OpenClaw host
openclaw plugins install @alavida-ai/outlook-plugin-openclaw
openclaw gateway restart
openclaw plugins inspect outlook
```

First-time setup requires running `outlook auth login` once on the OpenClaw host to populate the token cache. The plugin reads `~/.outlook-cli/tokens.json` directly.

The bundled skill (`packages/openclaw/skills/outlook/`) is shipped inside the plugin tarball via the `skills` field in `openclaw.plugin.json`.

## Supply-chain hardening

Per `~/.agentkb/alavida/wiki/javascript-supply-chain-security.md`:

- pnpm 11 pinned via `packageManager` in root `package.json`
- `preinstall` guard refuses `npm` / `yarn` / `bun` (`scripts/check-pnpm.mjs`)
- `pnpm-workspace.yaml`: per-package `allowBuilds`, `minimumReleaseAge: 1440`, `blockExoticSubdeps: true`, `verifyDepsBeforeRun: warn`
- CI uses `pnpm install --frozen-lockfile`

## Development

```bash
pnpm typecheck             # all 3 packages
pnpm test                  # vitest, ~117 tests
pnpm test:watch
pnpm lint
pnpm format
pnpm build
```

Tests live alongside source as `<name>.test.ts`. The token-cache, MSAL, and Graph tests use in-process fakes (no network).

## Why not `@microsoft/microsoft-graph-client` v3 → Kiota SDK

The classic `@microsoft/microsoft-graph-client` ships with one transitive dep (`tslib`), no install scripts, built-in retry middleware, and a `PageIterator` for `@odata.nextLink` cursoring. The newer Kiota-based `@microsoft/msgraph-sdk-javascript` is Microsoft's long-term direction but is still maturing and pulls in heavier deps. We use the classic SDK; switching is a future ticket.

## References

- Spec: [`docs/superpowers/specs/2026-05-28-outlook-typescript-rewrite-design.md`](docs/superpowers/specs/2026-05-28-outlook-typescript-rewrite-design.md)
- Foundation plan: [`docs/superpowers/plans/2026-05-28-outlook-foundation-and-whoami.md`](docs/superpowers/plans/2026-05-28-outlook-foundation-and-whoami.md)
- Granola precedent monorepo: https://github.com/alavida-ai/granola-plugin
- Alavida supply-chain playbook: `~/.agentkb/alavida/wiki/javascript-supply-chain-security.md`
- OpenClaw plugin distribution: `~/.agentkb/alavida/wiki/openclaw-plugin-distribution.md`
