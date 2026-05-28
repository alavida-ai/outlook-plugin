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
  stderr/stdout discipline. The one exception is auth — see §4 below, where we
  *do* invest in concrete robustness commitments the Python version lacks.
- OS keychain support for token storage. Native bindings add `allowBuilds`
  exposure and breakage risk; the file fallback we already rely on is sufficient
  and is the only option on the VPS anyway. Keychain stays open for v2 behind
  the `TokenCache` interface (§4.3).
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
    ├── core/        @alavida-ai/outlook-core              — Graph wrapper + MSAL
    ├── cli/         @alavida-ai/outlook-cli               — `outlook` binary
    └── openclaw/    @alavida-ai/outlook-plugin-openclaw   — OpenClaw plugin
```

**This shape is validated against the granola-plugin Python→TS migration**
(repo `alavida-ai/granola-plugin`, completed May 2026). Every structural
choice below — package names, `pnpm-workspace.yaml` block, `check-pnpm.mjs`
flavour, lazy-import CLI router, openclaw manifest shape — has a working
precedent in that repo and we copy from it rather than re-deriving.

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
| **`core`** (`@alavida-ai/outlook-core`) | MSAL device-code flow; token cache at `~/.outlook-cli/tokens.json` (0600); authenticated `graphFetch()`; typed message/event/folder/contact shapes; error mapping (`401 → AuthError`, `404 → NotFoundError`, `429/503 → throttled`). Pure library — no stdout/stderr, no `process.exit`. | `OutlookClient` with `.mail`, `.calendar`, `.contacts`, `.me`. Free functions: `loginDeviceCode()`, `logout()`, `status()`, `getAccessToken()`. |
| **`cli`** (`@alavida-ai/outlook-cli`) | Node's native `util.parseArgs` (no CLI-framework dep); `src/commands/<noun>-<verb>.ts` per command with `export function run(argv: string[]): Promise<number>`; lazy `await import('./commands/...')` from `src/index.ts`; `src/output.ts` for `printJson` / `eprintln` / `formatError`; stderr/stdout discipline (data → stdout, status/errors → stderr); `--json` flag; `--select` field projection; escape-decoding (`\n`, `\r`, `\t`, `\\` in `--body`/`--comment`). | The `outlook` binary plus per-command `--help`. No exports for downstream consumers. |
| **`openclaw`** (`@alavida-ai/outlook-plugin-openclaw`) | One `ToolDescriptor` per file in `src/tools/<tool-name>.ts`; central `registerTool` injects shared `output: 'pretty' \| 'json'` and `help: boolean` params and wraps in `withErrorMapping`; pretty renderers per tool. Depends on `@alavida-ai/outlook-core` via `workspace:*`. Package name follows the granola precedent (`@alavida-ai/<product>-plugin-openclaw`). | `default export = definePluginEntry({...})`. |

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

### 4.1 Why file-only (no OS keychain)

The Python implementation uses `keyring` (OS keychain) with a file fallback.
The TS rewrite is **file-only on purpose**:

- **Primary deployment is a headless VPS.** OpenClaw runs as a daemon under
  the chibote user. There is no OS keychain there; `keyring` already falls
  back to the file path on Linux without libsecret. The "primary path" is
  the fallback path.
- **Node OS-keychain bindings are a step backward.** `keytar` is unmaintained
  (2024 archival); `@napi-rs/keyring` is newer but adds native binary builds
  to `allowBuilds`, expands the install-script surface (the exact thing the
  KB's supply-chain hardening minimises), and breaks `pnpm install --frozen-lockfile`
  on novel architectures.
- **The "encryption at rest" win is mostly illusory on a developer laptop.**
  Same-user code can already read `$HOME`; keychain only protects against
  other UNIX users, which `chmod 0600` also handles. Full-disk encryption
  (FileVault / LUKS) is the real defence and doesn't depend on keychain.
- **Robustness is a different axis from encryption.** §4.2 below adds
  atomicity, locking, integrity checks, and a defined error taxonomy — none
  of which the current Python+keyring code has. Those are the wins for
  "very very robust auth".

OS keychain support stays open as a v2 optional-dependency feature; the
shape of the cache abstraction in §4.3 doesn't change if we add it.

### 4.2 Robustness commitments

Outlook auth is the load-bearing surface — a corrupted cache, a race between
concurrent refreshes, or a silent multi-account ambiguity all break the
plugin in ways that look like "the agent is broken". The TS rewrite makes
these concrete commitments, none of which exist (or are implicit-and-flaky)
in the current Python implementation.

1. **Atomic cache writes.** Every cache update writes to
   `tokens.json.tmp.<pid>.<rand>`, `fsync`s, then `fs.renameSync` to the
   canonical path. POSIX rename is atomic — a crash mid-write never leaves
   a half-written file. Wired into the MSAL `beforeCacheAccess` /
   `afterCacheAccess` plugin hooks so msal-node's serialiser benefits.

2. **Cross-process refresh lock.** Before any refresh, acquire a lock by
   creating `tokens.lock` with `fs.open(..., 'wx')` (`O_EXCL`). Retry with
   exponential backoff up to 30s. Stale-lock detection: if the lock file's
   mtime is older than 60s the holder is presumed dead and the lock is
   force-taken. CLI and OpenClaw plugin can run concurrently against the
   same cache without losing refresh tokens.

3. **Cache integrity check on read.** Validate the JSON parses and matches
   the expected MSAL cache schema (top-level keys: `AccessToken`,
   `RefreshToken`, `IdToken`, `Account`, `AppMetadata`). On corruption, log
   a structured warning to stderr (CLI) or emit a tool-error envelope
   (plugin) and treat as "not logged in" — **never** crash with
   `Unexpected end of JSON input`.

4. **Defined error taxonomy.** All subclass `AuthError`:
   - `AuthCacheMissingError` — no cache file; run `outlook auth login`
   - `AuthCacheCorruptError` — file present but unreadable; same fix
   - `AuthRefreshFailedError` — refresh token rejected; same fix
   - `AuthInteractionRequiredError` — Conditional Access / MFA re-prompt
     required; same fix
   - `AuthAmbiguousAccountError` — multiple accounts cached, none selected
   - `AuthLockTimeoutError` — couldn't acquire the refresh lock in 30s
   Every variant carries a human-readable `nextStep` string the CLI and
   plugin both surface verbatim.

5. **Multi-account handling.** MSAL's cache holds N accounts. Default
   behaviour: if exactly one account is cached, use it; if multiple, throw
   `AuthAmbiguousAccountError` listing the cached UPNs and require the
   caller to disambiguate via:
   - CLI: `--account <upn>` flag or `OUTLOOK_ACCOUNT` env var
   - Plugin: `account` field on the plugin's `configSchema`
   **Never silently pick `accounts[0]`.** The current Python does this —
   it's the cause of "wrong mailbox" bugs when a user has both personal
   and work Microsoft accounts.

6. **File permissions enforced on every write.** Parent directory `0700`,
   file `0600`. Re-applied after every rename; never trust prior state.
   Catches the case where a sibling process recreated the file with
   default umask, or a backup tool restored it with `0644`.

7. **Logout is idempotent and complete.** `outlook auth logout` removes
   `tokens.json`, the lock file, **and** any prior `keyring` entries the
   Python implementation may have left in macOS Keychain / Secret Service
   (migration kindness for users coming from the Python install — a
   one-shot best-effort delete that ignores not-found errors).

8. **Same Entra app id as the Python CLI.**
   `18f9e6ff-2b0a-423e-bb35-ab9b541e604e`, tenant `common`. Users
   migrating from Python don't re-consent. The `AZURE_CLIENT_ID` /
   `AZURE_TENANT_ID` override env vars keep working (escape hatch for
   clients on dedicated Entra apps).

9. **Same scope set as today.** `Mail.ReadWrite`, `Calendars.ReadWrite`,
   `Calendars.ReadWrite.Shared`, `Contacts.ReadWrite`, `User.Read`.
   `offline_access` added implicitly. Narrowing (e.g. `Mail.Read`-only
   variant) is a future ticket — would force re-consent.

10. **Auth test harness.** A fake transport mocks the MSAL discovery
    (`/.well-known/openid-configuration`), device-code, and token
    endpoints. Every error variant in §4.2.4 has at least one happy-path
    and one failure-path test. Concurrency test: two simultaneous refresh
    attempts against a single cache must both succeed and the cache must
    contain a valid refresh token at the end.

### 4.3 Cache abstraction

`core/src/auth/cache.ts` exports `TokenCache` interface with `load()`,
`save(blob)`, `clear()`, and `lock<T>(fn): Promise<T>`. The default
implementation is `FileTokenCache` (atomic writes + `O_EXCL` lock as above).
Any future backend (OS keychain, secret manager) plugs into the same
interface without touching call sites.

### 4.4 Threat model — explicit non-goals

- A malicious process running as the same UNIX user. The `0600` file
  permission only protects against *other* UNIX users; same-user code can
  read any file in `$HOME`. Defence is the OS user boundary, not us.
- A stolen, unlocked laptop. The cache is plaintext. FileVault / LUKS are
  the answer; we don't add a second encryption layer of dubious benefit.
- Microsoft revoking the embedded Entra app id. Fallback is the
  `AZURE_CLIENT_ID` override — same escape hatch as Python.
- Token sniffing in process memory. Out of scope; if the attacker is
  reading our process memory, plaintext tokens are the least of the
  user's problems.

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
| CLI framework | Node's `util.parseArgs` (no dep) | Granola-plugin precedent. Zero supply-chain surface; lazy command imports keep startup snappy. The CLI's command surface is shallow enough that we don't need yargs/commander/clipanion. |
| Tests | vitest | Same as sgil-crm. |
| Lint/format | eslint 9 + prettier | Same as sgil-crm. |
| Release | Changesets → path-tag-triggered GH Actions → GitHub Packages `@alavida-ai/*` | Per [openclaw-plugin-distribution.md](../../../.agentkb/alavida/wiki/openclaw-plugin-distribution.md). |
| Token cache | Plain file `~/.outlook-cli/tokens.json` (0600). No OS keychain. Atomic writes via tmpfile+rename; cross-process refresh lock via `O_EXCL` on `tokens.lock`; integrity check on read; defined error taxonomy. See §4.1–§4.4 for the full robustness story. | Native keychain bindings (keytar, @napi-rs/keyring) add `allowBuilds` entries + native-build complexity for a benefit (encryption at rest) that's mostly illusory once anything on the box can read the home directory. The Python code already falls back to a file; the VPS will use file-only too. The robustness wins this rewrite *does* care about (atomicity, locking, integrity, error taxonomy) are spec'd in §4.2. |
| Auth concurrency | Custom ~50-LOC `O_EXCL` file lock + tmpfile + rename | Zero deps. `proper-lockfile` was considered but pulls `graceful-fs` (a deep monkey-patcher of `fs`) — strictly more surface than the logic we'd write ourselves. |

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
- **v1.1 — `KeychainTokenCache` backend** behind the `TokenCache` interface
  (§4.3), gated by an `OUTLOOK_KEYCHAIN=1` env var (or CLI flag). Adds
  `@napi-rs/keyring` as an `optionalDependencies` entry — only platforms
  that actually use it pay the install cost. Trigger to prioritise: the
  first regulated-client deployment that asks "is the token cache
  encrypted at rest." File-only stays the default forever (the VPS has no
  keychain). See §4.1 for why this isn't a v1 blocker.
- A second `core` consumer (e.g. a small TUI mailbox triage tool).

## References

- [sgil-crm-plugin repo](https://github.com/alavida-ai/sgil-crm-plugin) — reference monorepo (4-package).
- [granola-plugin repo](https://github.com/alavida-ai/granola-plugin) — reference monorepo (3-package, the shape this design follows).
- [javascript-supply-chain-security.md](../../../.agentkb/alavida/wiki/javascript-supply-chain-security.md) — pnpm hardening playbook.
- [openclaw-plugin-distribution.md](../../../.agentkb/alavida/wiki/openclaw-plugin-distribution.md) — GH Packages distribution runbook.
- [openclaw-secrets.md](../../../.agentkb/alavida/wiki/openclaw-secrets.md) — config-shape rules for plugin secrets.
- [Current SKILL.md](../../../SKILL.md) — source of truth for the command surface the rewrite must hit parity with.
