# Outlook TypeScript rewrite — design

**Date:** 2026-05-28
**Owner:** chicote
**Status:** Approved (pending implementation)

## Goal

Rewrite the Python `outlook-cli` as a TypeScript pnpm monorepo so we can:

1. Ship a real OpenClaw plugin — one `ToolDescriptor` per existing CLI command — instead of having agents shell out to a `uv`-installed binary.
2. Keep a CLI for local testing and human use (Claude Code likes CLIs).
3. Apply the supply-chain hardening pattern Alavida has standardised on after the
   late-2025 npm attack wave (pnpm 11, `minimumReleaseAge`, `blockExoticSubdeps`,
   per-package `allowBuilds`, preinstall guard).
4. Use the same internal structure as `sgil-crm-plugin` so reviewers, CI, and
   future plugins benefit from a single mental model.

## Non-goals

- Adding new functionality. The new code must be **at parity** with the current
  Python CLI surface — same commands, same flags, same JSON envelopes, same
  stderr/stdout discipline.
- OS keychain support for token storage. Native bindings add `allowBuilds`
  exposure and breakage risk; the file fallback we already rely on is sufficient.
- Multi-tenant / per-client policy layers. Microsoft Graph's schema is the same
  for every tenant, so there is no equivalent of SGIL's `attio-sgil` policy
  package.
- Sending mail. The plugin remains draft-only (compliance posture, unchanged).

## Architecture

### Package layout (3 packages, mirrors granola-plugin)

```
outlook-cli/                                    # repo (unchanged GitHub URL)
├── package.json                                # workspace root
├── pnpm-workspace.yaml                         # supply-chain settings
├── scripts/check-pnpm.mjs                      # preinstall guard
├── tsconfig.base.json, tsconfig.json (composite refs)
├── eslint.config.js, .prettierrc.json, vitest.config.ts
├── .changeset/                                 # release management
├── .github/workflows/publish-openclaw.yml      # path-tag GH Packages publish
├── docs/superpowers/specs/                     # design docs (this file)
└── packages/
    ├── core/        @alavida-ai/outlook-core        — Graph wrapper + MSAL
    ├── cli/         @alavida-ai/outlook-cli         — `outlook` binary
    └── openclaw/    @alavida-ai/outlook-openclaw    — OpenClaw plugin
```

**Why 3 packages, not 4 (no sgil-equivalent business-policy layer).**
SGIL's `@alavida-ai/attio-sgil` exists because SGIL has client-specific Attio
collections, gated attributes, and policy decisions that don't generalise to
other Attio tenants. Outlook has no equivalent: Graph's surface is identical
for every tenant; there is no read-only-field policy, no collection gating, no
projection layer. The granola-plugin (also a thin wrapper around an external
API) is precedent for the 3-package shape.

### Package responsibilities

| Package | Owns | Public surface |
|---|---|---|
| **`core`** | MSAL device-code flow; token cache at `~/.outlook-cli/tokens.json` (0600); authenticated `graphFetch()`; typed message/event/folder/contact shapes; error mapping (`401 → AuthError`, `404 → NotFoundError`, `429/503 → throttled`). Pure library — no stdout/stderr, no process.exit. | `OutlookClient` with `.mail`, `.calendar`, `.contacts`, `.me`. Free functions: `loginDeviceCode()`, `logout()`, `status()`, `getAccessToken()`. |
| **`cli`** | Argument parsing (commander v12); pretty/json renderers; stderr/stdout discipline (data → stdout, status/errors → stderr); `--select` field projection; escape-decoding (`\n`, `\r`, `\t`, `\\` in `--body`/`--comment`); the `outlook` bin. | The `outlook` binary plus its `--help` text. No exports for downstream consumers. |
| **`openclaw`** | One `ToolDescriptor` per file in `src/tools/<tool-name>.ts`; central `registerTool` injects shared `output`/`help` params and wraps in `withErrorMapping`; pretty renderers per tool. Depends on `@alavida-ai/outlook-core` via `workspace:*`. | `default export = definePluginEntry({...})`. |

The OpenClaw plugin never shells out. The CLI and the plugin are two thin UIs
over the same `core`.

## Auth strategy

One token cache, two consumers.

- **CLI** runs `outlook auth login` → MSAL device-code → writes
  `~/.outlook-cli/tokens.json` (0600). First-line-of-stderr UX preserved
  (URL + code on the first line of stderr, blocks until sign-in completes,
  stdout stays clean for downstream piping).
- **OpenClaw plugin** reads the same file. `configSchema`:

  ```ts
  Type.Object({
    clientId:       Type.Optional(Type.String()),  // override embedded app id
    tenantId:       Type.Optional(Type.String()),  // default 'common'
    tokenCachePath: Type.Optional(Type.String()),  // default ~/.outlook-cli/tokens.json
  })
  ```

  No SecretRef shape needed (no API keys — auth is delegated through MSAL).
  String shorthand (`"${X}"`) suffices per
  [openclaw-secrets.md](../../../.agentkb/alavida/wiki/openclaw-secrets.md).

- **VPS first-time setup** requires `outlook auth login` to run once on the
  OpenClaw host. After that, refresh tokens renew silently — same lifecycle
  as today. If the cache is missing/expired, every tool throws a structured
  `AuthError` pointing the user at `outlook auth login`.

## OpenClaw tool surface

One tool per current CLI subcommand. Snake_case names, matching sgil-crm.

| Category | Tool names |
|---|---|
| **Identity (read)** | `whoami` |
| **Mail read** | `mail_list`, `mail_read`, `mail_search`, `mail_folders`, `mail_list_attachments`, `mail_download_attachment` |
| **Mail context-write (drafts, non-destructive triage)** | `mail_draft`, `mail_reply`, `mail_forward`, `mail_add_attachment`, `mail_mark`, `mail_flag`, `mail_importance` |
| **Mail state-change** | `mail_move`, `mail_delete` |
| **Calendar read** | `calendar_list`, `calendar_show`, `calendar_availability` |
| **Calendar write** | `calendar_create`, `calendar_update`, `calendar_delete`, `calendar_respond` |
| **Contacts** | `contacts_list` (stub — matches today's CLI stub) |

`auth login/logout/status` remain CLI-only — interactive device-code flow has no
home in a tool catalogue.

Every tool inherits shared `output: 'pretty' | 'json'` and `help: boolean`
params via the `registerTool` helper copied from sgil-crm.

## Tech stack

| Concern | Choice | Why |
|---|---|---|
| Runtime | Node ≥ 20 | OpenClaw SDK requirement; matches sgil-crm. |
| Package manager | pnpm 11.x (pinned via `packageManager`) | [javascript-supply-chain-security.md](../../../.agentkb/alavida/wiki/javascript-supply-chain-security.md). |
| MSAL | `@azure/msal-node` v3 | Official Microsoft device-code + token cache for Node. |
| Graph HTTP | Node 20 `fetch` + thin typed wrapper | Skips `@microsoft/microsoft-graph-client` — that lib brings its own auth-provider abstraction, large transitive surface, and types we don't need. Concise > comprehensive. |
| Validation | typebox | Same as sgil-crm `openclaw` package. |
| CLI framework | commander v12 | Smaller surface than yargs; well-used in the Azure CLI ecosystem. |
| Tests | vitest | Same as sgil-crm. |
| Lint/format | eslint 9 + prettier | Same as sgil-crm. |
| Release | Changesets → path-tag-triggered GH Actions → GitHub Packages `@alavida-ai/*` | Per [openclaw-plugin-distribution.md](../../../.agentkb/alavida/wiki/openclaw-plugin-distribution.md). |
| Token cache | Plain file `~/.outlook-cli/tokens.json` (0600). No OS keychain. | Native keychain bindings (keytar, @napi-rs/keyring) add `allowBuilds` entries + native build complexity for a benefit (encryption at rest) that's mostly illusory once anything on the box can read the home directory. The Python code already falls back to a file; the VPS will use file-only too. |

## Supply-chain hardening (verbatim per Alavida KB)

### Root `package.json`

```jsonc
{
  "packageManager": "pnpm@11.1.2",
  "engines": { "node": ">=20", "pnpm": ">=11" },
  "scripts": {
    "preinstall": "node scripts/check-pnpm.mjs",
    "build":     "pnpm -r run build",
    "typecheck": "pnpm -r run typecheck",
    "test":      "vitest run --passWithNoTests",
    "lint":      "eslint 'packages/**/*.ts' --no-error-on-unmatched-pattern",
    "format":    "prettier --write ."
  }
}
```

### `scripts/check-pnpm.mjs`

The 30-line user-agent check from the KB. Refuses npm/yarn/bun installs.

### `pnpm-workspace.yaml`

```yaml
packages:
  - 'packages/*'

allowBuilds:
  esbuild: true  # vitest dep
  # add as install surfaces them

minimumReleaseAge: 1440          # 24h cooldown
minimumReleaseAgeExclude: []
blockExoticSubdeps: true         # registry-only transitive deps
verifyDepsBeforeRun: warn        # composite refs false-positive on error
```

### CI

`pnpm install --frozen-lockfile` everywhere; `pnpm/action-setup@v6` +
`actions/setup-node@v4` with `cache: 'pnpm'`.

## Cutover plan

Branch off `main` as `typescript-rewrite` (no Linear ticket — small team,
no bookkeeping required for this rewrite).

1. **Commit 1 — Delete Python.** Remove `src/outlook_cli/`,
   `scripts/provision_entra_app.py`, `pyproject.toml`, `uv.lock`,
   `.python-version`, `.venv/`, `.ruff_cache/`, `dist/`, `PLAN.md`,
   `SUMMARY.md`. Git history preserves everything if we ever need it back.
   The `ALA-685` attachments commit (`71118c7`) sits on a separate branch,
   not main — when porting the mail-write slice the TS code reads the
   Python implementation from that branch as its parity reference for
   `mail_list_attachments`, `mail_download_attachment`,
   `mail_add_attachment`.
2. **Commit 2 — Workspace scaffold + supply-chain hardening.** Root
   `package.json`, `pnpm-workspace.yaml`, `scripts/check-pnpm.mjs`,
   tsconfig refs, eslint/prettier/vitest configs, empty package shells,
   `.changeset/`, GitHub Actions workflow.
3. **Commit 3 — `core` package skeleton.** MSAL device-code,
   token cache, `graphFetch` wrapper, `OutlookClient` with empty
   sub-clients, error classes, unit tests for the cache + error mapping.
4. **Commits 4–N — Vertical slices.** Each slice ports one feature area
   (auth → mail read → mail write → mail triage → calendar → contacts
   stub). Each slice lands the `core` method(s), the CLI sub-command, and
   the OpenClaw tool(s) **together**. CLI and plugin reach parity in
   lockstep; no "CLI ahead of plugin" or vice versa.
5. **Final commit — Skill + README + install docs.** Move
   `skills/outlook/SKILL.md` to `packages/openclaw/skills/outlook/`,
   rewrite the install section for GH Packages (`openclaw plugins install
   @alavida-ai/outlook-openclaw`), rewrite the README around the new layout.

Each vertical slice is a separate PR off the long-lived `typescript-rewrite`
branch, reviewed independently. The branch merges to main only after parity
is reached and the OpenClaw plugin has been smoke-tested on a real
OpenClaw host.

## Out of scope / future tickets

- `contacts_list` is a stub today; promoting to real Graph contacts is a
  follow-up.
- Widening `configSchema` to accept SecretRef object shape for `clientId`
  (per openclaw-secrets §4 recommendation). Not needed until we deploy in
  an environment that demands non-env secret resolution.
- OS keychain via a future optional dependency.
- A second `core` consumer (e.g. a small TUI mailbox triage tool).

## References

- [sgil-crm-plugin repo](https://github.com/alavida-ai/sgil-crm-plugin) — reference monorepo (4-package).
- [granola-plugin repo](https://github.com/alavida-ai/granola-plugin) — reference monorepo (3-package, the shape this design follows).
- [javascript-supply-chain-security.md](../../../.agentkb/alavida/wiki/javascript-supply-chain-security.md) — pnpm hardening playbook.
- [openclaw-plugin-distribution.md](../../../.agentkb/alavida/wiki/openclaw-plugin-distribution.md) — GH Packages distribution runbook.
- [openclaw-secrets.md](../../../.agentkb/alavida/wiki/openclaw-secrets.md) — config-shape rules for plugin secrets.
- [Current SKILL.md](../../../SKILL.md) — source of truth for the command surface the rewrite must hit parity with.
