# Outlook TypeScript rewrite — Foundation + whoami slice

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the Python `outlook-cli` with a TypeScript pnpm 11 monorepo and ship the first end-to-end working command — `outlook whoami` on the CLI and the `whoami` tool in the OpenClaw plugin — on top of MSAL device-code auth with the robustness commitments from the spec.

**Architecture:** 3-package pnpm workspace mirroring granola-plugin: `core` (MSAL + Graph), `cli` (terminal), `openclaw` (plugin). MSAL device-code → atomic file token cache → `@microsoft/microsoft-graph-client` → typed responses. See spec at `docs/superpowers/specs/2026-05-28-outlook-typescript-rewrite-design.md`.

**Tech Stack:** Node ≥ 20, pnpm 11.1.2, TypeScript 5.6, `@azure/msal-node` v3, `@microsoft/microsoft-graph-client`, `@microsoft/microsoft-graph-types`, typebox, vitest, eslint 9, prettier, changesets, openclaw plugin SDK.

**Branch:** `typescript-rewrite` (already cut off main). The spec is the only file on this branch beyond what's on main.

**Source-of-truth references:**
- Spec: `docs/superpowers/specs/2026-05-28-outlook-typescript-rewrite-design.md`
- Granola precedent: `/Users/alexandergarciachicote/code/projects/granola-plugin/`
- SGIL precedent: `/Users/alexandergarciachicote/code/projects/sgil-crm-plugin/`
- Python parity reference: `git show main:src/outlook_cli/` (read from main; deleted in Task 0.1)

---

## File Structure

By the end of this plan the repo looks like:

```
outlook-cli/
├── .changeset/config.json
├── .github/workflows/ci.yml              # typecheck + lint + test on PRs
├── .gitignore
├── .prettierignore
├── .prettierrc.json
├── docs/superpowers/specs/2026-05-28-outlook-typescript-rewrite-design.md   # exists already
├── docs/superpowers/plans/2026-05-28-outlook-foundation-and-whoami.md       # this file
├── eslint.config.js
├── package.json                           # workspace root
├── pnpm-workspace.yaml                    # supply-chain settings
├── scripts/check-pnpm.mjs                 # preinstall guard
├── tsconfig.base.json                     # shared TS config
├── tsconfig.json                          # composite refs
├── vitest.config.ts
└── packages/
    ├── core/
    │   ├── package.json
    │   ├── tsconfig.json
    │   └── src/
    │       ├── index.ts                   # public API barrel
    │       ├── auth/
    │       │   ├── cache.ts               # TokenCache interface + FileTokenCache
    │       │   ├── cache.test.ts
    │       │   ├── errors.ts              # AuthError taxonomy
    │       │   ├── errors.test.ts
    │       │   ├── msal.ts                # PublicClientApplication factory
    │       │   ├── device-code.ts         # loginDeviceCode()
    │       │   ├── silent.ts              # getAccessToken()
    │       │   ├── silent.test.ts
    │       │   ├── accounts.ts            # multi-account resolution
    │       │   ├── accounts.test.ts
    │       │   ├── logout.ts
    │       │   └── status.ts
    │       ├── graph/
    │       │   ├── client.ts              # makeGraphClient()
    │       │   ├── auth-provider.ts       # MSAL → Graph AuthProvider adapter
    │       │   ├── errors.ts              # Graph error mapping
    │       │   └── errors.test.ts
    │       ├── client.ts                  # OutlookClient (composes auth + graph + resources)
    │       ├── resources/
    │       │   └── me.ts                  # /me endpoint — only resource in this plan
    │       └── types.ts                   # exported entity types
    ├── cli/
    │   ├── package.json
    │   ├── tsconfig.json
    │   └── src/
    │       ├── index.ts                   # dispatcher + top-level --help
    │       ├── output.ts                  # printJson / eprintln / formatError
    │       ├── client.ts                  # shared OutlookClient construction
    │       └── commands/
    │           ├── auth-login.ts
    │           ├── auth-logout.ts
    │           ├── auth-status.ts
    │           └── whoami.ts
    └── openclaw/
        ├── package.json
        ├── tsconfig.json
        ├── openclaw.plugin.json
        ├── skills/outlook/SKILL.md        # placeholder — full move happens in last slice
        └── src/
            ├── index.ts                   # definePluginEntry + tool registration loop
            ├── client.ts                  # PluginConfig + getClient()
            ├── register.ts                # copied verbatim from granola
            ├── shared-schemas.ts          # copied verbatim from granola
            ├── errors.ts                  # withErrorMapping
            ├── pretty.ts                  # renderPretty dispatcher
            └── tools/
                └── whoami.ts
```

Subsequent plans add the `mail/`, `calendar/`, `contacts/` resource files and their CLI commands + openclaw tools without touching the foundation laid down here.

---

## Phase 0 — Repo prep (delete Python, scaffold workspace, supply-chain hardening)

### Task 0.1: Delete the Python tree

**Files:**
- Delete: `src/outlook_cli/`, `scripts/provision_entra_app.py`, `pyproject.toml`, `uv.lock`, `.python-version`, `.venv/`, `.ruff_cache/`, `dist/`, `PLAN.md`, `SUMMARY.md`

- [ ] **Step 1: Verify you're on the right branch**

```bash
git status
```

Expected: `On branch typescript-rewrite` with no uncommitted changes other than possibly `.claude/`. If you see anything else, stop and reconcile.

- [ ] **Step 2: Delete the Python tree**

```bash
git rm -r src/outlook_cli scripts/provision_entra_app.py pyproject.toml uv.lock .python-version PLAN.md SUMMARY.md
rm -rf .venv .ruff_cache dist
```

- [ ] **Step 3: Verify SKILL.md and README.md remain**

```bash
ls SKILL.md README.md
```

Expected: both files listed. (We rewrite them in the final slice; they stay as parity references in the meantime.)

- [ ] **Step 4: Commit**

```bash
git commit -m "$(cat <<'EOF'
Delete Python tree ahead of TypeScript rewrite

Per spec §Cutover plan, commit 1. Git history preserves the Python
implementation if we need to reference it. ALA-685 attachments work
(commit 71118c7) lives on its own branch and is referenced when we
port the mail-write slice.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 0.2: Add `scripts/check-pnpm.mjs` preinstall guard

**Files:**
- Create: `scripts/check-pnpm.mjs`

- [ ] **Step 1: Create the file**

Copy granola's version (more battle-tested than the KB's — permissive in CI and on empty user-agent, fail-closed only on a known-non-pnpm UA).

```js
/**
 * Refuse `npm install`, `yarn add`, `bun install` in this repo.
 *
 * Same idea as the `only-allow` npm package, inlined so it has zero
 * install-time dependencies (avoiding the dlx/npx user-agent clobber that
 * makes `npx only-allow pnpm` and `pnpm dlx only-allow pnpm` mis-fire under
 * pnpm 11).
 *
 * Permissive in two non-human contexts:
 *   1. CI ($CI truthy) — pnpm/action-setup doesn't always propagate
 *      npm_config_user_agent into preinstall subprocesses; CI's
 *      --frozen-lockfile is the real integrity check there.
 *   2. Empty npm_config_user_agent — ambiguous; we don't fail-closed since
 *      the consequence of a false-positive is a broken repo.
 *
 * We only fail when we KNOW the launching tool is npm, yarn, or bun.
 */

const CI = process.env.CI;
if (CI) {
  process.exit(0);
}

const userAgent = String(process.env.npm_config_user_agent || '');
if (!userAgent) {
  process.exit(0);
}

if (userAgent.startsWith('pnpm/')) {
  process.exit(0);
}

const detected = userAgent.split(/[\s/]/)[0] || 'unknown';
process.stderr.write(
  '\n' +
    '┌─────────────────────────────────────────────────────────────────┐\n' +
    '│  This repository requires pnpm.                                 │\n' +
    `│  Detected package manager: ${detected.padEnd(36)} │\n` +
    '│                                                                 │\n' +
    '│  Install pnpm via Corepack (uses the version pinned in this     │\n' +
    '│  repo via the `packageManager` field):                          │\n' +
    '│                                                                 │\n' +
    '│      corepack enable                                            │\n' +
    '│      pnpm install                                               │\n' +
    '│                                                                 │\n' +
    '│  Or install pnpm globally: npm install -g pnpm                  │\n' +
    '└─────────────────────────────────────────────────────────────────┘\n' +
    '\n',
);
process.exit(1);
```

- [ ] **Step 2: Verify it runs cleanly under pnpm context**

```bash
npm_config_user_agent='pnpm/11.1.2' node scripts/check-pnpm.mjs ; echo "exit=$?"
```

Expected: `exit=0`

- [ ] **Step 3: Verify it rejects npm**

```bash
npm_config_user_agent='npm/10.0.0 node/v20.0.0' node scripts/check-pnpm.mjs ; echo "exit=$?"
```

Expected: stderr message about pnpm, `exit=1`.

- [ ] **Step 4: Stage but DON'T commit yet (committed with the rest of the scaffold in Task 0.12)**

```bash
git add scripts/check-pnpm.mjs
```

---

### Task 0.3: Add root `.gitignore`

**Files:**
- Create: `.gitignore`

- [ ] **Step 1: Write the file**

```gitignore
node_modules/
dist/
*.tsbuildinfo
.DS_Store
coverage/
.env
.env.local
.vitest-cache/
```

- [ ] **Step 2: Stage**

```bash
git add .gitignore
```

---

### Task 0.4: Add root `package.json`

**Files:**
- Create: `package.json`

- [ ] **Step 1: Write the file**

```jsonc
{
  "name": "outlook-cli",
  "version": "0.0.0",
  "private": true,
  "description": "TypeScript monorepo for the Alavida Outlook plugin — CLI + OpenClaw + Claude Code skill for reading and triaging mail/calendar via Microsoft Graph.",
  "license": "UNLICENSED",
  "homepage": "https://github.com/alavida-ai/outlook-cli",
  "repository": {
    "type": "git",
    "url": "https://github.com/alavida-ai/outlook-cli.git"
  },
  "engines": {
    "node": ">=20",
    "pnpm": ">=11"
  },
  "packageManager": "pnpm@11.1.2",
  "scripts": {
    "preinstall": "node scripts/check-pnpm.mjs",
    "build": "pnpm -r run build",
    "typecheck": "pnpm -r run typecheck",
    "lint": "eslint 'packages/**/*.ts' --no-error-on-unmatched-pattern",
    "format": "prettier --write .",
    "test": "vitest run --passWithNoTests",
    "test:watch": "vitest",
    "changeset": "changeset",
    "version": "changeset version",
    "release": "pnpm build && changeset publish"
  },
  "devDependencies": {
    "@changesets/cli": "^2.27.9",
    "@types/node": "^20.16.10",
    "@eslint/js": "^9.12.0",
    "eslint": "^9.12.0",
    "typescript-eslint": "^8.8.0",
    "prettier": "^3.3.3",
    "typescript": "^5.6.2",
    "vitest": "^2.1.2"
  }
}
```

- [ ] **Step 2: Stage**

```bash
git add package.json
```

---

### Task 0.5: Add `pnpm-workspace.yaml` with supply-chain settings

**Files:**
- Create: `pnpm-workspace.yaml`

- [ ] **Step 1: Write the file**

Copy granola's verbatim — same allowBuilds set, same release-age, same lock policy.

```yaml
packages:
  - 'packages/*'

# ─────────────────────────────────────────────────────────────────────────────
# Supply-chain hardening — explicit even when these match pnpm 11 defaults.
# Defaults can change between minor versions; codifying them here means a
# future bump cannot silently weaken our posture without a diff in this file.
# Reference: ~/.agentkb/alavida/wiki/javascript-supply-chain-security.md
# ─────────────────────────────────────────────────────────────────────────────

# Per-package allowlist for preinstall/install/postinstall scripts.
# Every package that wants to run a script is enumerated with a deliberate
# decision. New transitive deps that try to run scripts will be blocked by
# default and pnpm will surface them on the next install.
#
# Why each:
#   esbuild           — downloads platform binary; vitest depends on it.
#   @google/genai     — Google Gemini SDK pulled in transitively via openclaw.
#                       Types only; no runtime Gemini call.
#   koffi             — FFI (native C bindings) pulled in via openclaw.
#   openclaw          — peer dep we type-check against; runtime is the gateway.
#   protobufjs        — pulled in transitively; not invoked at runtime.
#   tree-sitter-bash  — native bindings; pulled in transitively; not invoked.
#   sharp             — image processing native binary; pulled in transitively
#                       via openclaw; not invoked locally.
allowBuilds:
  esbuild: true
  '@google/genai': false
  koffi: false
  openclaw: false
  protobufjs: false
  tree-sitter-bash: false
  sharp: false

# Refuse package versions younger than 24 hours. Buys time for the community
# to detect and yank malicious releases (e.g. the TanStack incident, Sept 2025).
minimumReleaseAge: 1440
minimumReleaseAgeExclude: []

# Disallow transitive dependencies pulled from git URLs or arbitrary tarballs.
blockExoticSubdeps: true

# Verify node_modules matches lockfile before scripts. `warn` for composite
# refs; CI uses --frozen-lockfile as the hard check.
verifyDepsBeforeRun: warn
```

- [ ] **Step 2: Stage**

```bash
git add pnpm-workspace.yaml
```

---

### Task 0.6: Add `tsconfig.base.json` and root `tsconfig.json`

**Files:**
- Create: `tsconfig.base.json`
- Create: `tsconfig.json`

- [ ] **Step 1: Write `tsconfig.base.json`**

```jsonc
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "lib": ["ES2022"],
    "strict": true,
    "noImplicitAny": true,
    "noImplicitReturns": true,
    "noFallthroughCasesInSwitch": true,
    "noUncheckedIndexedAccess": true,
    "esModuleInterop": true,
    "isolatedModules": true,
    "forceConsistentCasingInFileNames": true,
    "skipLibCheck": true,
    "resolveJsonModule": true,
    "declaration": true,
    "declarationMap": true,
    "sourceMap": true,
    "composite": true,
    "incremental": true
  }
}
```

- [ ] **Step 2: Write root `tsconfig.json`**

```jsonc
{
  "files": [],
  "references": [
    { "path": "./packages/core" },
    { "path": "./packages/cli" },
    { "path": "./packages/openclaw" }
  ]
}
```

- [ ] **Step 3: Stage**

```bash
git add tsconfig.base.json tsconfig.json
```

---

### Task 0.7: Add `eslint.config.js` and Prettier configs

**Files:**
- Create: `eslint.config.js`
- Create: `.prettierrc.json`
- Create: `.prettierignore`

- [ ] **Step 1: Write `eslint.config.js`**

```js
// @ts-check
import eslint from '@eslint/js';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  {
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
    },
    rules: {
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
      '@typescript-eslint/consistent-type-imports': 'error',
    },
  },
  {
    ignores: ['**/dist/**', '**/node_modules/**', '**/*.tsbuildinfo'],
  },
);
```

- [ ] **Step 2: Write `.prettierrc.json`**

```json
{
  "semi": true,
  "singleQuote": true,
  "trailingComma": "all",
  "printWidth": 100,
  "tabWidth": 2,
  "arrowParens": "always"
}
```

- [ ] **Step 3: Write `.prettierignore`**

```
**/dist/**
**/node_modules/**
**/*.tsbuildinfo
pnpm-lock.yaml
```

- [ ] **Step 4: Stage**

```bash
git add eslint.config.js .prettierrc.json .prettierignore
```

---

### Task 0.8: Add `vitest.config.ts`

**Files:**
- Create: `vitest.config.ts`

- [ ] **Step 1: Write the file**

```ts
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    environment: 'node',
    include: ['packages/*/src/**/*.test.ts'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html'],
    },
  },
});
```

- [ ] **Step 2: Stage**

```bash
git add vitest.config.ts
```

---

### Task 0.9: Add `.changeset/config.json`

**Files:**
- Create: `.changeset/config.json`

- [ ] **Step 1: Write the file**

```json
{
  "$schema": "https://unpkg.com/@changesets/config@3.0.0/schema.json",
  "changelog": "@changesets/cli/changelog",
  "commit": false,
  "fixed": [],
  "linked": [],
  "access": "restricted",
  "baseBranch": "main",
  "updateInternalDependencies": "patch",
  "ignore": []
}
```

- [ ] **Step 2: Stage**

```bash
git add .changeset/config.json
```

---

### Task 0.10: Add CI workflow

**Files:**
- Create: `.github/workflows/ci.yml`

- [ ] **Step 1: Write the file**

```yaml
name: CI

on:
  pull_request:
  push:
    branches: [main, typescript-rewrite]

jobs:
  check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: pnpm/action-setup@v6
      - uses: actions/setup-node@v4
        with:
          node-version: '20'
          cache: 'pnpm'
      - run: pnpm install --frozen-lockfile
      - run: pnpm typecheck
      - run: pnpm lint
      - run: pnpm test
```

- [ ] **Step 2: Stage**

```bash
git add .github/workflows/ci.yml
```

---

### Task 0.11: Commit the scaffold

- [ ] **Step 1: Verify the staged set**

```bash
git status
```

Expected: 11 new files staged (`.changeset/config.json`, `.github/workflows/ci.yml`, `.gitignore`, `.prettierignore`, `.prettierrc.json`, `eslint.config.js`, `package.json`, `pnpm-workspace.yaml`, `scripts/check-pnpm.mjs`, `tsconfig.base.json`, `tsconfig.json`, `vitest.config.ts`). No deletes (those were committed in 0.1).

- [ ] **Step 2: Commit**

```bash
git commit -m "$(cat <<'EOF'
Scaffold pnpm 11 workspace + supply-chain hardening

Per spec §Cutover plan, commit 2. Layout copies granola-plugin
verbatim — same allowBuilds set, same release-age, same lock policy,
same eslint/prettier/vitest configs, same CI flow. The check-pnpm.mjs
guard is granola's CI-permissive variant, not the KB's stricter one.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 0.12: First `pnpm install` + verify guard

- [ ] **Step 1: Install**

```bash
pnpm install
```

Expected: a `pnpm-lock.yaml` is created, `node_modules/` populated. No warnings about unknown allowBuilds. Should complete in under a minute.

- [ ] **Step 2: Verify npm is refused**

```bash
npm install --dry-run 2>&1 | head -5
```

Expected: the boxed "This repository requires pnpm" message, exit code 1. (`--dry-run` skips lifecycle scripts so it *might* not trigger; if it doesn't, run plain `npm install` in a throwaway temp dir as documented in the KB.)

- [ ] **Step 3: Commit the lockfile**

```bash
git add pnpm-lock.yaml
git commit -m "$(cat <<'EOF'
Lockfile from first pnpm install

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 1 — `core` package: auth + token cache (the load-bearing surface)

### Task 1.1: Create `packages/core` scaffold

**Files:**
- Create: `packages/core/package.json`
- Create: `packages/core/tsconfig.json`
- Create: `packages/core/src/index.ts`

- [ ] **Step 1: Write `packages/core/package.json`**

```jsonc
{
  "name": "@alavida-ai/outlook-core",
  "version": "0.1.0",
  "description": "Microsoft Graph wrapper + MSAL device-code auth for Alavida's Outlook stack. Pure library — no stdout/stderr, no process.exit.",
  "type": "module",
  "license": "UNLICENSED",
  "main": "dist/index.js",
  "types": "dist/index.d.ts",
  "files": ["dist", "README.md"],
  "publishConfig": {
    "registry": "https://npm.pkg.github.com",
    "access": "restricted"
  },
  "scripts": {
    "build": "tsc -p tsconfig.json",
    "typecheck": "tsc -p tsconfig.json --noEmit"
  },
  "dependencies": {
    "@azure/msal-node": "^3.2.0",
    "@microsoft/microsoft-graph-client": "^3.0.7",
    "@microsoft/microsoft-graph-types": "^2.40.0"
  },
  "engines": { "node": ">=20" }
}
```

- [ ] **Step 2: Write `packages/core/tsconfig.json`**

```jsonc
{
  "extends": "../../tsconfig.base.json",
  "compilerOptions": {
    "rootDir": "src",
    "outDir": "dist",
    "tsBuildInfoFile": "tsconfig.tsbuildinfo"
  },
  "include": ["src/**/*.ts"],
  "exclude": ["src/**/*.test.ts"]
}
```

- [ ] **Step 3: Write `packages/core/src/index.ts`** (empty barrel for now)

```ts
// Public API of @alavida-ai/outlook-core. Exports added per task as the
// surface fills in. Anything not re-exported here is internal to the
// package.
export {};
```

- [ ] **Step 4: Run install + typecheck**

```bash
pnpm install
pnpm -F @alavida-ai/outlook-core typecheck
```

Expected: install adds `@azure/msal-node`, `@microsoft/microsoft-graph-client`, `@microsoft/microsoft-graph-types` to the lockfile; typecheck passes with no errors.

- [ ] **Step 5: Commit**

```bash
git add packages/core/package.json packages/core/tsconfig.json packages/core/src/index.ts pnpm-lock.yaml
git commit -m "$(cat <<'EOF'
core: scaffold @alavida-ai/outlook-core package

@azure/msal-node v3, @microsoft/microsoft-graph-client v3 (classic SDK),
@microsoft/microsoft-graph-types for type defs. Empty barrel index.ts —
real exports follow as the surface fills in.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.2: AuthError taxonomy

**Files:**
- Create: `packages/core/src/auth/errors.ts`
- Create: `packages/core/src/auth/errors.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/auth/errors.test.ts
import { describe, expect, it } from 'vitest';
import {
  AuthError,
  AuthCacheMissingError,
  AuthCacheCorruptError,
  AuthRefreshFailedError,
  AuthInteractionRequiredError,
  AuthAmbiguousAccountError,
  AuthLockTimeoutError,
} from './errors.js';

describe('AuthError taxonomy', () => {
  it('every variant is an instance of AuthError', () => {
    expect(new AuthCacheMissingError()).toBeInstanceOf(AuthError);
    expect(new AuthCacheCorruptError('bad json')).toBeInstanceOf(AuthError);
    expect(new AuthRefreshFailedError('refresh rejected')).toBeInstanceOf(AuthError);
    expect(new AuthInteractionRequiredError('MFA required')).toBeInstanceOf(AuthError);
    expect(new AuthAmbiguousAccountError(['a@x.com', 'b@y.com'])).toBeInstanceOf(AuthError);
    expect(new AuthLockTimeoutError(30_000)).toBeInstanceOf(AuthError);
  });

  it('every variant carries a nextStep string', () => {
    const variants: AuthError[] = [
      new AuthCacheMissingError(),
      new AuthCacheCorruptError('x'),
      new AuthRefreshFailedError('x'),
      new AuthInteractionRequiredError('x'),
      new AuthAmbiguousAccountError(['a@x.com']),
      new AuthLockTimeoutError(30_000),
    ];
    for (const e of variants) {
      expect(e.nextStep).toMatch(/outlook auth login|--account|wait/);
    }
  });

  it('AuthAmbiguousAccountError lists UPNs in the message', () => {
    const e = new AuthAmbiguousAccountError(['alice@example.com', 'bob@example.com']);
    expect(e.message).toContain('alice@example.com');
    expect(e.message).toContain('bob@example.com');
    expect(e.accounts).toEqual(['alice@example.com', 'bob@example.com']);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/errors.test.ts
```

Expected: FAIL — `Cannot find module './errors.js'`.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/auth/errors.ts

/**
 * Base class for every auth-time failure raised by `@alavida-ai/outlook-core`.
 *
 * Tools (CLI and OpenClaw) catch `AuthError` and surface `.nextStep` verbatim
 * to the human. Subclasses carry typed context the formatter can use to print
 * something more specific.
 */
export class AuthError extends Error {
  /** Short, user-facing remediation. Always populated. */
  public readonly nextStep: string;

  constructor(message: string, nextStep: string) {
    super(message);
    this.name = new.target.name;
    this.nextStep = nextStep;
  }
}

export class AuthCacheMissingError extends AuthError {
  constructor() {
    super('No cached Microsoft account.', 'Run `outlook auth login` to sign in.');
  }
}

export class AuthCacheCorruptError extends AuthError {
  constructor(reason: string) {
    super(
      `Token cache is unreadable (${reason}).`,
      'Run `outlook auth logout` then `outlook auth login` to start fresh.',
    );
  }
}

export class AuthRefreshFailedError extends AuthError {
  constructor(reason: string) {
    super(
      `Silent token refresh failed: ${reason}.`,
      'Run `outlook auth login` to re-authenticate.',
    );
  }
}

export class AuthInteractionRequiredError extends AuthError {
  constructor(reason: string) {
    super(
      `Microsoft requires interactive sign-in (${reason}).`,
      'Run `outlook auth login` to complete sign-in.',
    );
  }
}

export class AuthAmbiguousAccountError extends AuthError {
  /** UPNs of the cached accounts, in cache order. */
  public readonly accounts: readonly string[];

  constructor(accounts: readonly string[]) {
    super(
      `Multiple accounts cached (${accounts.join(', ')}); none selected.`,
      'Pass `--account <upn>` (CLI), set `OUTLOOK_ACCOUNT=<upn>`, or set the `account` config field (plugin).',
    );
    this.accounts = accounts;
  }
}

export class AuthLockTimeoutError extends AuthError {
  constructor(timeoutMs: number) {
    super(
      `Couldn't acquire token-cache refresh lock within ${timeoutMs} ms.`,
      'Another process is mid-refresh; wait a few seconds and retry, or delete `~/.outlook-cli/tokens.lock` if no other process is using it.',
    );
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/errors.test.ts
```

Expected: 3 tests pass.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/auth/errors.ts packages/core/src/auth/errors.test.ts
git commit -m "$(cat <<'EOF'
core: define AuthError taxonomy

Six variants subclass AuthError, each with a typed shape and a
human-readable nextStep. Tests assert instanceof, presence of nextStep,
and that AuthAmbiguousAccountError surfaces the cached UPNs.

Per spec §4.2.4.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.3: `TokenCache` interface + in-memory test double

**Files:**
- Create: `packages/core/src/auth/cache.ts` (interface only for now)
- Create: `packages/core/src/auth/cache.test.ts` (in-memory test double, exercises the contract)

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/auth/cache.test.ts
import { describe, expect, it } from 'vitest';
import { type TokenCache, InMemoryTokenCache } from './cache.js';

describe('TokenCache contract', () => {
  it('round-trips a string blob', async () => {
    const cache: TokenCache = new InMemoryTokenCache();
    expect(await cache.load()).toBe(null);
    await cache.save('hello');
    expect(await cache.load()).toBe('hello');
  });

  it('clear() leaves load() returning null', async () => {
    const cache: TokenCache = new InMemoryTokenCache();
    await cache.save('hello');
    await cache.clear();
    expect(await cache.load()).toBe(null);
  });

  it('lock() serializes critical sections', async () => {
    const cache: TokenCache = new InMemoryTokenCache();
    const order: string[] = [];
    const a = cache.lock(async () => {
      order.push('a-start');
      await new Promise((r) => setTimeout(r, 20));
      order.push('a-end');
      return 'a';
    });
    const b = cache.lock(async () => {
      order.push('b-start');
      order.push('b-end');
      return 'b';
    });
    const [ra, rb] = await Promise.all([a, b]);
    expect(ra).toBe('a');
    expect(rb).toBe('b');
    expect(order).toEqual(['a-start', 'a-end', 'b-start', 'b-end']);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/cache.test.ts
```

Expected: FAIL — `Cannot find module './cache.js'`.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/auth/cache.ts

/**
 * Storage backend for the MSAL serialised token cache.
 *
 * Two production backends are envisioned:
 *   - `FileTokenCache` — default. Atomic writes via tmpfile+rename;
 *     cross-process refresh lock via `O_EXCL` on a sibling `.lock` file.
 *     See `./cache-file.ts` (added in Task 1.4).
 *   - `KeychainTokenCache` — v1.1 (out of scope here). Same interface;
 *     swaps the storage primitive.
 *
 * The `InMemoryTokenCache` in this file is a test double — it satisfies
 * the contract with an in-process map + a single-slot promise lock.
 */
export interface TokenCache {
  /** Returns the cached MSAL serialised blob, or `null` if absent / cleared. */
  load(): Promise<string | null>;

  /** Atomically write `blob` to the cache. Overwrites any prior content. */
  save(blob: string): Promise<void>;

  /** Remove the cache entirely. Idempotent. */
  clear(): Promise<void>;

  /**
   * Run `fn` while holding an exclusive lock on the cache.
   *
   * The lock is process-local for in-memory implementations and cross-process
   * (POSIX `O_EXCL` file lock) for `FileTokenCache`. The caller does NOT need
   * to call `load`/`save` inside `fn`; the lock is purely for serialising
   * read-modify-write cycles around MSAL's silent-refresh logic.
   */
  lock<T>(fn: () => Promise<T>): Promise<T>;
}

/** In-memory backend used by tests. Not exported from the package index. */
export class InMemoryTokenCache implements TokenCache {
  private blob: string | null = null;
  private chain: Promise<unknown> = Promise.resolve();

  async load(): Promise<string | null> {
    return this.blob;
  }

  async save(blob: string): Promise<void> {
    this.blob = blob;
  }

  async clear(): Promise<void> {
    this.blob = null;
  }

  async lock<T>(fn: () => Promise<T>): Promise<T> {
    const previous = this.chain;
    let release!: () => void;
    const next = new Promise<void>((resolve) => {
      release = resolve;
    });
    this.chain = previous.then(() => next);
    await previous;
    try {
      return await fn();
    } finally {
      release();
    }
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/cache.test.ts
```

Expected: 3 tests pass.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/auth/cache.ts packages/core/src/auth/cache.test.ts
git commit -m "$(cat <<'EOF'
core: TokenCache interface + InMemoryTokenCache test double

Per spec §4.3. The contract is load/save/clear/lock; the in-memory
implementation satisfies it with a promise chain for serialisation.
FileTokenCache lands in the next task.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.4: `FileTokenCache` — atomic write + read with integrity check

**Files:**
- Create: `packages/core/src/auth/cache-file.ts`
- Create: `packages/core/src/auth/cache-file.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/auth/cache-file.test.ts
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { FileTokenCache } from './cache-file.js';
import { AuthCacheCorruptError } from './errors.js';

describe('FileTokenCache', () => {
  let dir: string;
  let cache: FileTokenCache;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cache-'));
    cache = new FileTokenCache(join(dir, 'tokens.json'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('returns null when the file does not exist', async () => {
    expect(await cache.load()).toBe(null);
  });

  it('round-trips an MSAL-shaped JSON blob', async () => {
    const blob = JSON.stringify({
      AccessToken: {},
      RefreshToken: {},
      IdToken: {},
      Account: {},
      AppMetadata: {},
    });
    await cache.save(blob);
    expect(await cache.load()).toBe(blob);
  });

  it('writes with 0600 file mode and 0700 parent dir mode', async () => {
    const blob = '{"AccessToken":{},"RefreshToken":{},"IdToken":{},"Account":{},"AppMetadata":{}}';
    await cache.save(blob);
    const fileStat = await stat(join(dir, 'tokens.json'));
    expect(fileStat.mode & 0o777).toBe(0o600);
    const dirStat = await stat(dir);
    // Parent dir mode may vary depending on mkdtemp default; cache should
    // tighten the dir it owns (a fresh subdir under `dir`, set explicitly
    // when the file is first written into a missing tree).
    const cacheDirCache = new FileTokenCache(join(dir, 'sub', 'tokens.json'));
    await cacheDirCache.save(blob);
    const subStat = await stat(join(dir, 'sub'));
    expect(subStat.mode & 0o777).toBe(0o700);
  });

  it('clear() removes the file and is idempotent', async () => {
    await cache.save('{"AccessToken":{},"RefreshToken":{},"IdToken":{},"Account":{},"AppMetadata":{}}');
    await cache.clear();
    expect(await cache.load()).toBe(null);
    await cache.clear(); // idempotent
    expect(await cache.load()).toBe(null);
  });

  it('treats non-JSON as corrupt', async () => {
    await writeFile(join(dir, 'tokens.json'), 'not json{');
    await expect(cache.load()).rejects.toBeInstanceOf(AuthCacheCorruptError);
  });

  it('treats JSON missing MSAL top-level keys as corrupt', async () => {
    await writeFile(join(dir, 'tokens.json'), JSON.stringify({ random: 'thing' }));
    await expect(cache.load()).rejects.toBeInstanceOf(AuthCacheCorruptError);
  });

  it('atomic write: a poisoned tmp file does not affect the canonical path', async () => {
    const goodBlob = '{"AccessToken":{},"RefreshToken":{},"IdToken":{},"Account":{},"AppMetadata":{}}';
    await cache.save(goodBlob);
    // Drop a half-written tmp file alongside the canonical one.
    await writeFile(join(dir, 'tokens.json.tmp.99999.deadbeef'), 'half-writ');
    expect(await cache.load()).toBe(goodBlob);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/cache-file.test.ts
```

Expected: FAIL — `Cannot find module './cache-file.js'`.

- [ ] **Step 3: Write the implementation (atomic write + integrity, not lock yet)**

```ts
// packages/core/src/auth/cache-file.ts
import { randomBytes } from 'node:crypto';
import {
  chmod,
  mkdir,
  open,
  readFile,
  rename,
  rm,
  stat,
  unlink,
} from 'node:fs/promises';
import { dirname } from 'node:path';

import { AuthCacheCorruptError } from './errors.js';
import type { TokenCache } from './cache.js';

const REQUIRED_KEYS = ['AccessToken', 'RefreshToken', 'IdToken', 'Account', 'AppMetadata'] as const;

/**
 * Default `TokenCache` backend: a single JSON file at `path`, written
 * atomically (`tmpfile → fsync → rename`) with 0600 permissions and a
 * 0700 parent directory.
 *
 * Cross-process serialisation lives in Task 1.5 (`lock()` method). Until
 * then this implementation throws `Unsupported` on `lock()` calls so the
 * file-write tests can run cleanly.
 */
export class FileTokenCache implements TokenCache {
  constructor(public readonly path: string) {}

  async load(): Promise<string | null> {
    let raw: string;
    try {
      raw = await readFile(this.path, 'utf8');
    } catch (err) {
      if (isENOENT(err)) return null;
      throw err;
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch (err) {
      throw new AuthCacheCorruptError(`invalid JSON: ${(err as Error).message}`);
    }
    if (!isMsalShape(parsed)) {
      throw new AuthCacheCorruptError('missing MSAL top-level keys');
    }
    return raw;
  }

  async save(blob: string): Promise<void> {
    const dir = dirname(this.path);
    await mkdir(dir, { recursive: true, mode: 0o700 });
    await chmod(dir, 0o700).catch(() => {}); // best-effort tighten

    const tmp = `${this.path}.tmp.${process.pid}.${randomBytes(4).toString('hex')}`;
    const handle = await open(tmp, 'w', 0o600);
    try {
      await handle.writeFile(blob);
      await handle.sync();
    } finally {
      await handle.close();
    }
    await rename(tmp, this.path);
    await chmod(this.path, 0o600);
  }

  async clear(): Promise<void> {
    await rm(this.path, { force: true });
  }

  // Real implementation lands in Task 1.5.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  async lock<T>(_fn: () => Promise<T>): Promise<T> {
    throw new Error('FileTokenCache.lock() not implemented yet (Task 1.5).');
  }
}

function isENOENT(err: unknown): boolean {
  return typeof err === 'object' && err !== null && (err as { code?: string }).code === 'ENOENT';
}

function isMsalShape(value: unknown): boolean {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;
  return REQUIRED_KEYS.every((k) => k in v);
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/cache-file.test.ts
```

Expected: 7 tests pass. (The `lock()` placeholder is not exercised here.)

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/auth/cache-file.ts packages/core/src/auth/cache-file.test.ts
git commit -m "$(cat <<'EOF'
core: FileTokenCache — atomic writes + integrity check

Per spec §4.2.{1,3,6}:
  - Atomic write: tmpfile → writeFile → fsync → close → rename
  - 0600 file, 0700 parent dir, re-applied on every save
  - load() throws AuthCacheCorruptError on bad JSON or non-MSAL shape
  - clear() is idempotent

Cross-process lock lands in the next task; lock() throws until then.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.5: `FileTokenCache.lock()` — cross-process `O_EXCL` lock

**Files:**
- Modify: `packages/core/src/auth/cache-file.ts` (implement `lock()`)
- Create: `packages/core/src/auth/cache-file-lock.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/auth/cache-file-lock.test.ts
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { FileTokenCache } from './cache-file.js';
import { AuthLockTimeoutError } from './errors.js';

describe('FileTokenCache.lock()', () => {
  let dir: string;
  let cache: FileTokenCache;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-lock-'));
    cache = new FileTokenCache(join(dir, 'tokens.json'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('runs the critical section and releases the lock on success', async () => {
    const result = await cache.lock(async () => 'ok');
    expect(result).toBe('ok');
    // Lock file removed:
    await expect(cache.lock(async () => 'again')).resolves.toBe('again');
  });

  it('releases the lock when the critical section throws', async () => {
    await expect(cache.lock(async () => { throw new Error('boom'); })).rejects.toThrow('boom');
    await expect(cache.lock(async () => 'after')).resolves.toBe('after');
  });

  it('serializes concurrent same-process calls', async () => {
    const events: string[] = [];
    const a = cache.lock(async () => {
      events.push('a-start');
      await new Promise((r) => setTimeout(r, 25));
      events.push('a-end');
    });
    const b = cache.lock(async () => {
      events.push('b-start');
      events.push('b-end');
    });
    await Promise.all([a, b]);
    expect(events).toEqual(['a-start', 'a-end', 'b-start', 'b-end']);
  });

  it('throws AuthLockTimeoutError when the lock is held past the timeout', async () => {
    // Acquire and hold:
    let release!: () => void;
    const held = cache.lock(async () => {
      await new Promise<void>((r) => { release = r; });
    });
    // Override the default timeout for the test:
    const fast = new FileTokenCache(join(dir, 'tokens.json'));
    await expect(fast.lock(async () => 'never', { timeoutMs: 200 })).rejects.toBeInstanceOf(
      AuthLockTimeoutError,
    );
    release();
    await held;
  });

  it('force-takes a stale lock (mtime > maxAge)', async () => {
    // Drop a stale lock file with an old mtime:
    const lockPath = join(dir, 'tokens.lock');
    await writeFile(lockPath, '0');
    const { utimesSync } = await import('node:fs');
    const old = Date.now() / 1000 - 120; // 2 minutes ago
    utimesSync(lockPath, old, old);
    await expect(cache.lock(async () => 'taken', { maxLockAgeMs: 60_000 })).resolves.toBe(
      'taken',
    );
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/cache-file-lock.test.ts
```

Expected: FAIL — `lock() not implemented yet`.

- [ ] **Step 3: Implement `lock()` and widen the method signature**

In `packages/core/src/auth/cache-file.ts`, replace the placeholder `lock` with:

```ts
import { setTimeout as sleep } from 'node:timers/promises';
import { writeFile, utimes } from 'node:fs/promises';
import { openSync, closeSync } from 'node:fs';
```

Add a `lockPath` property and the new method:

```ts
export interface LockOptions {
  /** Total time to keep retrying lock acquisition. Default 30s. */
  timeoutMs?: number;
  /**
   * If an existing lock file is older than this, presume the holder is dead
   * and force-take it. Default 60s.
   */
  maxLockAgeMs?: number;
}

const DEFAULT_TIMEOUT_MS = 30_000;
const DEFAULT_MAX_LOCK_AGE_MS = 60_000;

// inside class FileTokenCache:

get lockPath(): string {
  return `${this.path}.lock`;
}

async lock<T>(fn: () => Promise<T>, opts: LockOptions = {}): Promise<T> {
  const timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const maxLockAgeMs = opts.maxLockAgeMs ?? DEFAULT_MAX_LOCK_AGE_MS;
  const start = Date.now();
  let attempt = 0;

  while (true) {
    try {
      const fd = openSync(this.lockPath, 'wx', 0o600);
      try {
        await writeFile(this.lockPath, String(process.pid));
      } finally {
        closeSync(fd);
      }
      break; // lock acquired
    } catch (err) {
      if (!isEEXIST(err)) throw err;
      // Lock exists. Check staleness.
      try {
        const st = await stat(this.lockPath);
        const age = Date.now() - st.mtimeMs;
        if (age > maxLockAgeMs) {
          await unlink(this.lockPath).catch(() => {});
          continue; // retry immediately
        }
      } catch (statErr) {
        if (isENOENT(statErr)) continue; // gone between EEXIST and stat — retry
        throw statErr;
      }
      if (Date.now() - start > timeoutMs) {
        throw new AuthLockTimeoutError(timeoutMs);
      }
      const backoff = Math.min(250 * 2 ** attempt, 1_000);
      attempt += 1;
      await sleep(backoff);
    }
  }

  try {
    return await fn();
  } finally {
    await unlink(this.lockPath).catch(() => {});
  }
}
```

And add the helper at the bottom of the file:

```ts
function isEEXIST(err: unknown): boolean {
  return typeof err === 'object' && err !== null && (err as { code?: string }).code === 'EEXIST';
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/cache-file-lock.test.ts
```

Expected: 5 tests pass. Same-process concurrency test serializes correctly; timeout test rejects with `AuthLockTimeoutError`; stale-lock test force-takes.

- [ ] **Step 5: Re-run the full cache-file suite for regression**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/cache-file.test.ts src/auth/cache-file-lock.test.ts
```

Expected: 12 tests pass total.

- [ ] **Step 6: Commit**

```bash
git add packages/core/src/auth/cache-file.ts packages/core/src/auth/cache-file-lock.test.ts
git commit -m "$(cat <<'EOF'
core: FileTokenCache.lock() — O_EXCL cross-process lock

Per spec §4.2.2. Exponential backoff up to 30s timeout. Stale-lock
detection: if the lock file is older than 60s the holder is presumed
dead and the lock is force-taken. Same-process concurrency is
serialised because acquire-retry-loop happens on every call.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.6: MSAL `PublicClientApplication` factory + cache plugin

**Files:**
- Create: `packages/core/src/auth/msal.ts`
- Create: `packages/core/src/auth/msal.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/auth/msal.test.ts
import { describe, expect, it } from 'vitest';

import { InMemoryTokenCache } from './cache.js';
import { buildMsalApp, EMBEDDED_CLIENT_ID, EMBEDDED_TENANT, OUTLOOK_SCOPES } from './msal.js';

describe('buildMsalApp', () => {
  it('returns a PublicClientApplication wired to the embedded app id by default', () => {
    const app = buildMsalApp({ cache: new InMemoryTokenCache() });
    expect(app).toBeDefined();
    expect(EMBEDDED_CLIENT_ID).toBe('18f9e6ff-2b0a-423e-bb35-ab9b541e604e');
    expect(EMBEDDED_TENANT).toBe('common');
  });

  it('accepts clientId and tenantId overrides', () => {
    const app = buildMsalApp({
      cache: new InMemoryTokenCache(),
      clientId: 'other-id',
      tenantId: 'other-tenant',
    });
    expect(app).toBeDefined();
  });

  it('exports the spec scope set', () => {
    expect(OUTLOOK_SCOPES).toEqual([
      'Mail.ReadWrite',
      'Calendars.ReadWrite',
      'Calendars.ReadWrite.Shared',
      'Contacts.ReadWrite',
      'User.Read',
    ]);
  });

  it('reads cache via the plugin hook on first token-cache access', async () => {
    const cache = new InMemoryTokenCache();
    await cache.save('{"AccessToken":{},"RefreshToken":{},"IdToken":{},"Account":{},"AppMetadata":{}}');
    const app = buildMsalApp({ cache });
    // Forcing a cache read via the public surface:
    const accounts = await app.getTokenCache().getAllAccounts();
    expect(accounts).toEqual([]);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/msal.test.ts
```

Expected: FAIL — `Cannot find module './msal.js'`.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/auth/msal.ts
import {
  PublicClientApplication,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from '@azure/msal-node';

import type { TokenCache } from './cache.js';

/**
 * The shared multi-tenant Entra app that ships with `outlook-cli`. Same id
 * the Python implementation used, so users migrating from Python don't
 * re-consent.
 */
export const EMBEDDED_CLIENT_ID = '18f9e6ff-2b0a-423e-bb35-ab9b541e604e';

/**
 * `common` accepts both personal Microsoft accounts and any work/school
 * tenant. MSAL resolves the actual tenant from the user's sign-in. Required
 * for multi-tenant apps.
 */
export const EMBEDDED_TENANT = 'common';

/**
 * Delegated scopes the CLI requests at sign-in. `offline_access` is added
 * implicitly by MSAL for public clients.
 */
export const OUTLOOK_SCOPES = [
  'Mail.ReadWrite',
  'Calendars.ReadWrite',
  'Calendars.ReadWrite.Shared',
  'Contacts.ReadWrite',
  'User.Read',
] as const;

export interface BuildMsalAppOptions {
  cache: TokenCache;
  clientId?: string;
  tenantId?: string;
}

/**
 * Construct an MSAL `PublicClientApplication` wired to our `TokenCache`.
 *
 * MSAL serialises its cache to a string blob; we treat that blob as opaque
 * and let `TokenCache` own atomicity, integrity, and locking (see §4.2).
 */
export function buildMsalApp(options: BuildMsalAppOptions): PublicClientApplication {
  const clientId = options.clientId ?? EMBEDDED_CLIENT_ID;
  const tenantId = options.tenantId ?? EMBEDDED_TENANT;

  const cachePlugin: ICachePlugin = {
    async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
      const blob = await options.cache.load().catch((err) => {
        // Corrupt cache: surface as "empty" to MSAL; the AuthCacheCorruptError
        // is re-raised by the upper-layer getAccessToken on the next user op.
        ctx.tokenCache.deserialize('');
        throw err;
      });
      if (blob) ctx.tokenCache.deserialize(blob);
    },
    async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
      if (!ctx.cacheHasChanged) return;
      await options.cache.save(ctx.tokenCache.serialize());
    },
  };

  const config: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: { cachePlugin },
  };
  return new PublicClientApplication(config);
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/msal.test.ts
```

Expected: 4 tests pass.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/auth/msal.ts packages/core/src/auth/msal.test.ts
git commit -m "$(cat <<'EOF'
core: MSAL PublicClientApplication factory + cache plugin

Per spec §4.2.{1,8,9}. Embedded client id and scope set match the
Python implementation so migrating users skip re-consent. cachePlugin
delegates serialisation atomicity and integrity to TokenCache.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.7: Multi-account resolution helper

**Files:**
- Create: `packages/core/src/auth/accounts.ts`
- Create: `packages/core/src/auth/accounts.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/auth/accounts.test.ts
import { describe, expect, it } from 'vitest';
import type { AccountInfo } from '@azure/msal-node';

import { resolveAccount } from './accounts.js';
import { AuthAmbiguousAccountError, AuthCacheMissingError } from './errors.js';

const a = (username: string): AccountInfo => ({
  homeAccountId: `${username}-home`,
  environment: 'login.microsoftonline.com',
  tenantId: 'tenant-id',
  username,
  localAccountId: `${username}-local`,
});

describe('resolveAccount', () => {
  it('throws AuthCacheMissingError when no accounts cached', () => {
    expect(() => resolveAccount([], undefined)).toThrow(AuthCacheMissingError);
  });

  it('returns the single cached account when there is exactly one', () => {
    const acc = a('only@example.com');
    expect(resolveAccount([acc], undefined)).toBe(acc);
  });

  it('throws AuthAmbiguousAccountError when multiple accounts and no preference', () => {
    expect(() => resolveAccount([a('one@x.com'), a('two@y.com')], undefined)).toThrow(
      AuthAmbiguousAccountError,
    );
  });

  it('selects by UPN when preference is provided', () => {
    const accounts = [a('one@x.com'), a('two@y.com')];
    expect(resolveAccount(accounts, 'two@y.com')).toBe(accounts[1]);
  });

  it('UPN matching is case-insensitive', () => {
    const accounts = [a('Alice@Example.com')];
    expect(resolveAccount(accounts, 'alice@EXAMPLE.com')).toBe(accounts[0]);
  });

  it('throws AuthAmbiguousAccountError listing UPNs when preference does not match', () => {
    const accounts = [a('one@x.com'), a('two@y.com')];
    let caught: unknown;
    try {
      resolveAccount(accounts, 'three@z.com');
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(AuthAmbiguousAccountError);
    const e = caught as AuthAmbiguousAccountError;
    expect(e.accounts).toEqual(['one@x.com', 'two@y.com']);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/accounts.test.ts
```

Expected: FAIL — module not found.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/auth/accounts.ts
import type { AccountInfo } from '@azure/msal-node';

import { AuthAmbiguousAccountError, AuthCacheMissingError } from './errors.js';

/**
 * Pick a single cached account, applying the caller's preference if any.
 *
 * Per spec §4.2.5 — **never** silently pick `accounts[0]`. If the cache
 * holds more than one account and the caller hasn't disambiguated,
 * `AuthAmbiguousAccountError` lists the cached UPNs for the human/agent
 * to choose from.
 *
 * @param accounts MSAL's cached `AccountInfo[]` in cache order.
 * @param preferredUpn UPN to select. Undefined = "I have no preference".
 *                     Case-insensitive match.
 */
export function resolveAccount(
  accounts: readonly AccountInfo[],
  preferredUpn: string | undefined,
): AccountInfo {
  if (accounts.length === 0) {
    throw new AuthCacheMissingError();
  }

  if (preferredUpn !== undefined) {
    const needle = preferredUpn.toLowerCase();
    const hit = accounts.find((a) => a.username.toLowerCase() === needle);
    if (hit) return hit;
    throw new AuthAmbiguousAccountError(accounts.map((a) => a.username));
  }

  if (accounts.length === 1) {
    return accounts[0]!;
  }

  throw new AuthAmbiguousAccountError(accounts.map((a) => a.username));
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/accounts.test.ts
```

Expected: 6 tests pass.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/auth/accounts.ts packages/core/src/auth/accounts.test.ts
git commit -m "$(cat <<'EOF'
core: resolveAccount helper — never silently picks accounts[0]

Per spec §4.2.5. Single-account: use it. Multi-account + no preference:
AuthAmbiguousAccountError listing UPNs. Multi-account + preference:
case-insensitive UPN match or AuthAmbiguousAccountError. Zero-account:
AuthCacheMissingError.

This closes the "wrong mailbox" bug class from the Python implementation.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.8: `getAccessToken()` — silent refresh with the cache lock

**Files:**
- Create: `packages/core/src/auth/silent.ts`
- Create: `packages/core/src/auth/silent.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/auth/silent.test.ts
import { describe, expect, it, vi } from 'vitest';
import type { AccountInfo, PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from './cache.js';
import { getAccessToken } from './silent.js';
import {
  AuthInteractionRequiredError,
  AuthRefreshFailedError,
} from './errors.js';

const account: AccountInfo = {
  homeAccountId: 'home',
  environment: 'login.microsoftonline.com',
  tenantId: 't',
  username: 'a@x.com',
  localAccountId: 'local',
};

function makeApp(opts: {
  accounts: AccountInfo[];
  acquireSilentResult: { accessToken: string; expiresOn: Date } | Error;
}): PublicClientApplication {
  return {
    getTokenCache: () => ({
      getAllAccounts: async () => opts.accounts,
    }),
    acquireTokenSilent: vi.fn(async () => {
      if (opts.acquireSilentResult instanceof Error) throw opts.acquireSilentResult;
      return opts.acquireSilentResult;
    }),
  } as unknown as PublicClientApplication;
}

describe('getAccessToken', () => {
  it('returns the token from acquireTokenSilent on the happy path', async () => {
    const cache = new InMemoryTokenCache();
    const expiresOn = new Date(Date.now() + 3_600_000);
    const app = makeApp({ accounts: [account], acquireSilentResult: { accessToken: 'tok', expiresOn } });
    const result = await getAccessToken({ app, cache, preferredUpn: undefined });
    expect(result.accessToken).toBe('tok');
    expect(result.expiresOn).toBe(expiresOn);
    expect(result.account).toBe(account);
  });

  it('throws AuthInteractionRequiredError when MSAL says interaction required', async () => {
    const cache = new InMemoryTokenCache();
    const err = Object.assign(new Error('interaction_required'), {
      errorCode: 'interaction_required',
    });
    const app = makeApp({ accounts: [account], acquireSilentResult: err });
    await expect(
      getAccessToken({ app, cache, preferredUpn: undefined }),
    ).rejects.toBeInstanceOf(AuthInteractionRequiredError);
  });

  it('throws AuthRefreshFailedError on any other MSAL error', async () => {
    const cache = new InMemoryTokenCache();
    const app = makeApp({
      accounts: [account],
      acquireSilentResult: new Error('network down'),
    });
    await expect(
      getAccessToken({ app, cache, preferredUpn: undefined }),
    ).rejects.toBeInstanceOf(AuthRefreshFailedError);
  });

  it('runs the refresh under cache.lock()', async () => {
    const cache = new InMemoryTokenCache();
    const lockSpy = vi.spyOn(cache, 'lock');
    const app = makeApp({
      accounts: [account],
      acquireSilentResult: { accessToken: 'tok', expiresOn: new Date(Date.now() + 1_000) },
    });
    await getAccessToken({ app, cache, preferredUpn: undefined });
    expect(lockSpy).toHaveBeenCalledTimes(1);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/silent.test.ts
```

Expected: FAIL — module not found.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/auth/silent.ts
import type { AccountInfo, PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import { resolveAccount } from './accounts.js';
import {
  AuthInteractionRequiredError,
  AuthRefreshFailedError,
} from './errors.js';
import { OUTLOOK_SCOPES } from './msal.js';

export interface GetAccessTokenInput {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn: string | undefined;
}

export interface AccessTokenResult {
  accessToken: string;
  expiresOn: Date;
  account: AccountInfo;
}

/**
 * Acquire a valid Graph access token for the cached account.
 *
 * Holds the cache lock for the entire silent-refresh cycle so concurrent
 * CLI + plugin processes don't race against each other (spec §4.2.2). All
 * `AuthError` variants the caller might see are typed.
 */
export async function getAccessToken(input: GetAccessTokenInput): Promise<AccessTokenResult> {
  const { app, cache, preferredUpn } = input;
  return cache.lock(async () => {
    const accounts = await app.getTokenCache().getAllAccounts();
    const account = resolveAccount(accounts, preferredUpn);
    try {
      const result = await app.acquireTokenSilent({
        account,
        scopes: [...OUTLOOK_SCOPES],
      });
      if (!result?.accessToken || !result.expiresOn) {
        throw new AuthRefreshFailedError('MSAL returned no access token');
      }
      return {
        accessToken: result.accessToken,
        expiresOn: result.expiresOn,
        account,
      };
    } catch (err) {
      if (err instanceof AuthRefreshFailedError) throw err;
      if (isInteractionRequired(err)) {
        throw new AuthInteractionRequiredError((err as Error).message ?? 'interaction required');
      }
      throw new AuthRefreshFailedError(toMessage(err));
    }
  });
}

function isInteractionRequired(err: unknown): boolean {
  if (typeof err !== 'object' || err === null) return false;
  const code = (err as { errorCode?: string }).errorCode;
  return code === 'interaction_required' || code === 'consent_required' || code === 'login_required';
}

function toMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  return String(err);
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/silent.test.ts
```

Expected: 4 tests pass.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/auth/silent.ts packages/core/src/auth/silent.test.ts
git commit -m "$(cat <<'EOF'
core: getAccessToken — silent refresh under the cache lock

Per spec §4.2.{2,4}. Resolves the account via resolveAccount() (throws
AuthCacheMissingError / AuthAmbiguousAccountError), runs
acquireTokenSilent inside cache.lock(), maps MSAL's interaction_required
family to AuthInteractionRequiredError and anything else to
AuthRefreshFailedError.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.9: `loginDeviceCode()` — interactive device-code flow

**Files:**
- Create: `packages/core/src/auth/device-code.ts`
- Create: `packages/core/src/auth/device-code.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/auth/device-code.test.ts
import { describe, expect, it, vi } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from './cache.js';
import { loginDeviceCode } from './device-code.js';

describe('loginDeviceCode', () => {
  it('forwards the device-code message via the callback and resolves with the result', async () => {
    const cache = new InMemoryTokenCache();
    const messages: string[] = [];
    const acquireMock = vi.fn(async (req: any) => {
      req.deviceCodeCallback({
        message: 'To sign in, use a web browser to open https://microsoft.com/devicelogin and enter the code AB12CD34',
        userCode: 'AB12CD34',
        verificationUri: 'https://microsoft.com/devicelogin',
        deviceCode: 'd',
        expiresIn: 900,
        interval: 5,
      });
      return {
        accessToken: 'tok',
        account: {
          homeAccountId: 'h',
          environment: 'e',
          tenantId: 't',
          username: 'a@x.com',
          localAccountId: 'l',
        },
        expiresOn: new Date(Date.now() + 3_600_000),
      };
    });
    const app = { acquireTokenByDeviceCode: acquireMock } as unknown as PublicClientApplication;

    const result = await loginDeviceCode({
      app,
      cache,
      onDeviceCode: (msg) => messages.push(msg.message),
    });

    expect(result.account.username).toBe('a@x.com');
    expect(messages).toHaveLength(1);
    expect(messages[0]).toContain('AB12CD34');
  });

  it('runs under cache.lock()', async () => {
    const cache = new InMemoryTokenCache();
    const lockSpy = vi.spyOn(cache, 'lock');
    const app = {
      acquireTokenByDeviceCode: vi.fn(async () => ({
        accessToken: 'tok',
        account: {
          homeAccountId: 'h',
          environment: 'e',
          tenantId: 't',
          username: 'a@x.com',
          localAccountId: 'l',
        },
        expiresOn: new Date(Date.now() + 1_000),
      })),
    } as unknown as PublicClientApplication;
    await loginDeviceCode({ app, cache, onDeviceCode: () => undefined });
    expect(lockSpy).toHaveBeenCalledTimes(1);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/device-code.test.ts
```

Expected: FAIL — module not found.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/auth/device-code.ts
import type { AccountInfo, DeviceCodeResponse, PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import { AuthRefreshFailedError } from './errors.js';
import { OUTLOOK_SCOPES } from './msal.js';

export interface LoginDeviceCodeInput {
  app: PublicClientApplication;
  cache: TokenCache;
  /**
   * Called once when MSAL has the device-code response. The CLI prints
   * `info.message` to stderr verbatim — that text contains the URL and
   * the user code Microsoft expects you to display.
   */
  onDeviceCode: (info: DeviceCodeResponse) => void;
}

export interface LoginResult {
  account: AccountInfo;
  expiresOn: Date;
}

/**
 * Run the device-code flow under the cache lock. Blocks (up to MSAL's default
 * polling window — ~15 min) until the user finishes sign-in in their browser,
 * then writes the resulting tokens through the cache plugin.
 */
export async function loginDeviceCode(input: LoginDeviceCodeInput): Promise<LoginResult> {
  const { app, cache, onDeviceCode } = input;
  return cache.lock(async () => {
    const result = await app.acquireTokenByDeviceCode({
      scopes: [...OUTLOOK_SCOPES],
      deviceCodeCallback: (info) => onDeviceCode(info),
    });
    if (!result?.account || !result.expiresOn) {
      throw new AuthRefreshFailedError('device-code flow returned no account');
    }
    return { account: result.account, expiresOn: result.expiresOn };
  });
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/device-code.test.ts
```

Expected: 2 tests pass.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/auth/device-code.ts packages/core/src/auth/device-code.test.ts
git commit -m "$(cat <<'EOF'
core: loginDeviceCode — interactive device-code flow

Per spec §4. Acquires under cache.lock() so a parallel silent refresh
on the OpenClaw plugin doesn't race against the human signing in.
deviceCodeCallback is forwarded verbatim to the caller — the CLI prints
the message to stderr to preserve the existing first-line-of-stderr UX.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.10: `logout()` and `status()`

**Files:**
- Create: `packages/core/src/auth/logout.ts`
- Create: `packages/core/src/auth/logout.test.ts`
- Create: `packages/core/src/auth/status.ts`
- Create: `packages/core/src/auth/status.test.ts`

- [ ] **Step 1: Write the failing tests**

```ts
// packages/core/src/auth/logout.test.ts
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { FileTokenCache } from './cache-file.js';
import { logout } from './logout.js';

describe('logout', () => {
  let dir: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-logout-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('removes the token cache file', async () => {
    const cache = new FileTokenCache(join(dir, 'tokens.json'));
    await cache.save('{"AccessToken":{},"RefreshToken":{},"IdToken":{},"Account":{},"AppMetadata":{}}');
    await logout({ cache });
    expect(existsSync(join(dir, 'tokens.json'))).toBe(false);
  });

  it('removes the lock file if present', async () => {
    const cache = new FileTokenCache(join(dir, 'tokens.json'));
    await writeFile(join(dir, 'tokens.json.lock'), '0');
    await logout({ cache });
    expect(existsSync(join(dir, 'tokens.json.lock'))).toBe(false);
  });

  it('is idempotent when nothing is cached', async () => {
    const cache = new FileTokenCache(join(dir, 'tokens.json'));
    await logout({ cache });
    await logout({ cache });
  });
});
```

```ts
// packages/core/src/auth/status.test.ts
import { describe, expect, it } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from './cache.js';
import { status } from './status.js';
import { AuthAmbiguousAccountError } from './errors.js';

function appWith(usernames: string[]): PublicClientApplication {
  return {
    getTokenCache: () => ({
      getAllAccounts: async () =>
        usernames.map((u) => ({
          homeAccountId: `${u}-h`,
          environment: 'e',
          tenantId: 't',
          username: u,
          localAccountId: 'l',
        })),
    }),
  } as unknown as PublicClientApplication;
}

describe('status', () => {
  it('returns null when no accounts are cached', async () => {
    expect(await status({ app: appWith([]), cache: new InMemoryTokenCache() })).toBe(null);
  });

  it('returns the single cached account when there is one', async () => {
    const result = await status({ app: appWith(['a@x.com']), cache: new InMemoryTokenCache() });
    expect(result?.username).toBe('a@x.com');
  });

  it('throws AuthAmbiguousAccountError when multiple accounts and no preference', async () => {
    await expect(
      status({ app: appWith(['a@x.com', 'b@y.com']), cache: new InMemoryTokenCache() }),
    ).rejects.toBeInstanceOf(AuthAmbiguousAccountError);
  });

  it('honours preferredUpn when provided', async () => {
    const result = await status({
      app: appWith(['a@x.com', 'b@y.com']),
      cache: new InMemoryTokenCache(),
      preferredUpn: 'b@y.com',
    });
    expect(result?.username).toBe('b@y.com');
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/logout.test.ts src/auth/status.test.ts
```

Expected: both FAIL — modules not found.

- [ ] **Step 3: Write the implementations**

```ts
// packages/core/src/auth/logout.ts
import { unlink } from 'node:fs/promises';

import type { TokenCache } from './cache.js';
import type { FileTokenCache } from './cache-file.js';

export interface LogoutInput {
  cache: TokenCache;
}

/**
 * Tear down all auth state owned by this package. Idempotent.
 *
 * Per spec §4.2.7, this also makes a best-effort attempt to delete the
 * `keyring` entry the Python CLI may have left around (migration kindness).
 * The deletion is wrapped in a try/catch — we never crash logout because
 * the user is on a system without libsecret/Keychain.
 */
export async function logout(input: LogoutInput): Promise<void> {
  const { cache } = input;
  await cache.clear();

  // FileTokenCache exposes lockPath; for other backends this is a no-op.
  const lockPath = (cache as Partial<FileTokenCache>).lockPath;
  if (typeof lockPath === 'string') {
    await unlink(lockPath).catch(() => {});
  }
  // Best-effort: the Python CLI used keyring service "outlook-cli" / key "default".
  // We don't take a keytar dep just for migration; the file deletion above is
  // enough on systems where Python tokens were file-cached. Keychain users
  // can run `security delete-generic-password -s outlook-cli -a default` once.
}
```

```ts
// packages/core/src/auth/status.ts
import type { AccountInfo, PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import { resolveAccount } from './accounts.js';
import { AuthAmbiguousAccountError } from './errors.js';

export interface StatusInput {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn?: string;
}

/**
 * Return the cached account info if logged in, `null` if not.
 *
 * Multi-account ambiguity still throws — `outlook auth status` is a check,
 * not a workaround for picking among accounts. The caller passes
 * `preferredUpn` when they already know which mailbox they're asking about.
 */
export async function status(input: StatusInput): Promise<AccountInfo | null> {
  const { app, cache: _cache, preferredUpn } = input;
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length === 0) return null;
  try {
    return resolveAccount(accounts, preferredUpn);
  } catch (err) {
    if (err instanceof AuthAmbiguousAccountError) throw err;
    return null;
  }
}
```

- [ ] **Step 4: Run the tests to verify they pass**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/auth/logout.test.ts src/auth/status.test.ts
```

Expected: 3 + 4 = 7 tests pass.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/auth/logout.ts packages/core/src/auth/logout.test.ts packages/core/src/auth/status.ts packages/core/src/auth/status.test.ts
git commit -m "$(cat <<'EOF'
core: logout() and status() helpers

logout: idempotent, removes tokens.json + lock file. Notes the Python
keyring migration step inline rather than taking a keytar dep just for
the one-shot cleanup.

status: returns the cached AccountInfo or null. Multi-account ambiguity
still throws — status is a check, not a tiebreaker.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 1.11: Wire core public auth API + barrel exports

**Files:**
- Modify: `packages/core/src/index.ts`

- [ ] **Step 1: Update the barrel**

```ts
// packages/core/src/index.ts
export {
  AuthError,
  AuthCacheMissingError,
  AuthCacheCorruptError,
  AuthRefreshFailedError,
  AuthInteractionRequiredError,
  AuthAmbiguousAccountError,
  AuthLockTimeoutError,
} from './auth/errors.js';

export type { TokenCache } from './auth/cache.js';
export { FileTokenCache } from './auth/cache-file.js';
export type { LockOptions } from './auth/cache-file.js';

export {
  buildMsalApp,
  EMBEDDED_CLIENT_ID,
  EMBEDDED_TENANT,
  OUTLOOK_SCOPES,
} from './auth/msal.js';
export type { BuildMsalAppOptions } from './auth/msal.js';

export { resolveAccount } from './auth/accounts.js';

export { getAccessToken } from './auth/silent.js';
export type { AccessTokenResult, GetAccessTokenInput } from './auth/silent.js';

export { loginDeviceCode } from './auth/device-code.js';
export type { LoginDeviceCodeInput, LoginResult } from './auth/device-code.js';

export { logout } from './auth/logout.js';
export type { LogoutInput } from './auth/logout.js';

export { status } from './auth/status.js';
export type { StatusInput } from './auth/status.js';
```

- [ ] **Step 2: Run typecheck + full test suite**

```bash
pnpm -F @alavida-ai/outlook-core typecheck && pnpm test
```

Expected: typecheck passes; ~33 tests pass.

- [ ] **Step 3: Commit**

```bash
git add packages/core/src/index.ts
git commit -m "$(cat <<'EOF'
core: public barrel — auth surface

Re-exports the error taxonomy, TokenCache + FileTokenCache, MSAL
factory, account resolution, getAccessToken, loginDeviceCode, logout,
status. Anything not listed here is internal to @alavida-ai/outlook-core.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 2 — `core` package: Graph client wiring

### Task 2.1: MSAL → Graph `AuthenticationProvider` adapter

**Files:**
- Create: `packages/core/src/graph/auth-provider.ts`
- Create: `packages/core/src/graph/auth-provider.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/graph/auth-provider.test.ts
import { describe, expect, it, vi } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from '../auth/cache.js';
import { MsalAuthenticationProvider } from './auth-provider.js';

describe('MsalAuthenticationProvider', () => {
  it('returns the access token from getAccessToken()', async () => {
    const cache = new InMemoryTokenCache();
    const app = {
      getTokenCache: () => ({
        getAllAccounts: async () => [
          { homeAccountId: 'h', environment: 'e', tenantId: 't', username: 'a@x.com', localAccountId: 'l' },
        ],
      }),
      acquireTokenSilent: vi.fn(async () => ({
        accessToken: 'graph-tok',
        expiresOn: new Date(Date.now() + 3_600_000),
      })),
    } as unknown as PublicClientApplication;

    const provider = new MsalAuthenticationProvider({ app, cache, preferredUpn: undefined });
    expect(await provider.getAccessToken()).toBe('graph-tok');
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/graph/auth-provider.test.ts
```

Expected: FAIL — module not found.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/graph/auth-provider.ts
import type { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import type { PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from '../auth/cache.js';
import { getAccessToken } from '../auth/silent.js';

export interface MsalAuthenticationProviderOptions {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn: string | undefined;
}

/**
 * `AuthenticationProvider` for `@microsoft/microsoft-graph-client` that
 * delegates to our MSAL silent-refresh path. The Graph SDK calls
 * `getAccessToken()` once per request (modulo its own caching).
 */
export class MsalAuthenticationProvider implements AuthenticationProvider {
  constructor(private readonly options: MsalAuthenticationProviderOptions) {}

  async getAccessToken(): Promise<string> {
    const result = await getAccessToken(this.options);
    return result.accessToken;
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/graph/auth-provider.test.ts
```

Expected: 1 test passes.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/graph/auth-provider.ts packages/core/src/graph/auth-provider.test.ts
git commit -m "$(cat <<'EOF'
core/graph: MSAL → AuthenticationProvider adapter

Implements @microsoft/microsoft-graph-client's AuthenticationProvider
interface by delegating to our existing silent.getAccessToken(). The
SDK calls getAccessToken() once per request; we already own caching
and locking inside silent.ts.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 2.2: Graph error mapping

**Files:**
- Create: `packages/core/src/graph/errors.ts`
- Create: `packages/core/src/graph/errors.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/graph/errors.test.ts
import { describe, expect, it } from 'vitest';
import { GraphError } from '@microsoft/microsoft-graph-client';

import {
  CoreError,
  NotFoundError,
  ThrottledError,
  ServerError,
  NetworkError,
  liftGraphError,
} from './errors.js';
import { AuthError } from '../auth/errors.js';

function makeGraphError(statusCode: number, message = 'msg'): GraphError {
  const err = new GraphError(statusCode, message);
  return err;
}

describe('liftGraphError', () => {
  it('maps 401 to AuthError', () => {
    expect(liftGraphError(makeGraphError(401))).toBeInstanceOf(AuthError);
  });

  it('maps 404 to NotFoundError', () => {
    expect(liftGraphError(makeGraphError(404))).toBeInstanceOf(NotFoundError);
  });

  it('maps 429 to ThrottledError with retryAfter when available', () => {
    const err = makeGraphError(429, 'throttled');
    (err as unknown as { headers?: Record<string, string> }).headers = { 'retry-after': '17' };
    const lifted = liftGraphError(err) as ThrottledError;
    expect(lifted).toBeInstanceOf(ThrottledError);
    expect(lifted.retryAfterSeconds).toBe(17);
  });

  it('maps 503 to ServerError', () => {
    expect(liftGraphError(makeGraphError(503))).toBeInstanceOf(ServerError);
  });

  it('wraps unknown thrown values as NetworkError', () => {
    expect(liftGraphError(new TypeError('fetch failed'))).toBeInstanceOf(NetworkError);
  });

  it('passes through non-Error thrown values as NetworkError', () => {
    expect(liftGraphError('weird')).toBeInstanceOf(NetworkError);
  });

  it('CoreError is the common base', () => {
    expect(liftGraphError(makeGraphError(500))).toBeInstanceOf(CoreError);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/graph/errors.test.ts
```

Expected: FAIL — module not found.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/graph/errors.ts
import { GraphError } from '@microsoft/microsoft-graph-client';

import { AuthError, AuthRefreshFailedError } from '../auth/errors.js';

/** Base for every non-auth runtime error from `core`. */
export class CoreError extends Error {
  constructor(message: string) {
    super(message);
    this.name = new.target.name;
  }
}

export class NotFoundError extends CoreError {}
export class ServerError extends CoreError {}
export class NetworkError extends CoreError {}

export class ThrottledError extends CoreError {
  constructor(message: string, public readonly retryAfterSeconds: number | null) {
    super(message);
  }
}

/**
 * Lift any error thrown by the Graph SDK into our taxonomy. Callers wrap
 * every `client.api(...).get/post/etc()` with this so consumers never see a
 * bare `GraphError`.
 */
export function liftGraphError(err: unknown): Error {
  if (err instanceof GraphError) {
    const status = err.statusCode;
    if (status === 401) return new AuthRefreshFailedError(err.message || 'unauthorised');
    if (status === 404) return new NotFoundError(err.message || 'not found');
    if (status === 429) {
      const headers = (err as unknown as { headers?: Record<string, string> }).headers ?? {};
      const ra = parseInt(headers['retry-after'] ?? '', 10);
      return new ThrottledError(err.message || 'throttled', Number.isFinite(ra) ? ra : null);
    }
    if (status >= 500) return new ServerError(err.message || `graph ${status}`);
    return new CoreError(`graph ${status}: ${err.message ?? ''}`.trim());
  }
  if (err instanceof AuthError) return err;
  if (err instanceof Error) return new NetworkError(err.message || 'network error');
  return new NetworkError(String(err));
}
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/graph/errors.test.ts
```

Expected: 7 tests pass.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/graph/errors.ts packages/core/src/graph/errors.test.ts
git commit -m "$(cat <<'EOF'
core/graph: lift GraphError into our taxonomy

Per spec §Package responsibilities for core. The SDK's GraphError plus
arbitrary thrown values become CoreError subclasses: NotFoundError (404),
ThrottledError (429, with retry-after), ServerError (5xx), NetworkError
(non-Error throws and TypeErrors from fetch). 401 maps to
AuthRefreshFailedError so the auth UX surfaces.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 2.3: `makeGraphClient` helper

**Files:**
- Create: `packages/core/src/graph/client.ts`

- [ ] **Step 1: Write the implementation (smoke-tested via `me.get()` next task)**

```ts
// packages/core/src/graph/client.ts
import { Client } from '@microsoft/microsoft-graph-client';
import type { PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from '../auth/cache.js';
import { MsalAuthenticationProvider } from './auth-provider.js';

export interface MakeGraphClientOptions {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn: string | undefined;
}

/**
 * Construct a configured `@microsoft/microsoft-graph-client` instance whose
 * auth provider delegates to our MSAL silent-refresh path.
 *
 * The SDK's middleware chain handles `Retry-After` on 429 / 503 out of the
 * box; we don't add custom middleware here.
 */
export function makeGraphClient(options: MakeGraphClientOptions): Client {
  const authProvider = new MsalAuthenticationProvider(options);
  return Client.initWithMiddleware({ authProvider });
}
```

- [ ] **Step 2: Verify the file typechecks**

```bash
pnpm -F @alavida-ai/outlook-core typecheck
```

Expected: passes.

- [ ] **Step 3: Commit**

```bash
git add packages/core/src/graph/client.ts
git commit -m "$(cat <<'EOF'
core/graph: makeGraphClient helper

Client.initWithMiddleware activates the SDK's default middleware chain
(retry, redirect, telemetry). Auth provider delegates to our MSAL
silent-refresh path.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 2.4: `OutlookClient` + `me.get()` (the only resource in this plan)

**Files:**
- Create: `packages/core/src/resources/me.ts`
- Create: `packages/core/src/client.ts`
- Create: `packages/core/src/client.test.ts`
- Modify: `packages/core/src/index.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/client.test.ts
import { describe, expect, it, vi } from 'vitest';
import { Client } from '@microsoft/microsoft-graph-client';

import { OutlookClient } from './client.js';

function fakeGraphClient(meResponse: unknown): Client {
  return {
    api: vi.fn(() => ({
      get: vi.fn(async () => meResponse),
    })),
  } as unknown as Client;
}

describe('OutlookClient.me.get', () => {
  it('returns the Graph /me payload as-is', async () => {
    const sample = {
      id: 'user-id',
      displayName: 'Alice Example',
      mail: 'alice@example.com',
      userPrincipalName: 'alice@example.com',
      jobTitle: 'Engineer',
      department: 'Eng',
      officeLocation: 'London',
    };
    const client = new OutlookClient(fakeGraphClient(sample));
    expect(await client.me.get()).toEqual(sample);
  });

  it('passes /me to .api()', async () => {
    const graph = fakeGraphClient({});
    const client = new OutlookClient(graph);
    await client.me.get();
    expect(graph.api).toHaveBeenCalledWith('/me');
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/client.test.ts
```

Expected: FAIL — module not found.

- [ ] **Step 3: Write the resource and client**

```ts
// packages/core/src/resources/me.ts
import type { Client } from '@microsoft/microsoft-graph-client';
import type { User } from '@microsoft/microsoft-graph-types';

import { liftGraphError } from '../graph/errors.js';

/** Subset of `User` we surface today. The full Graph type is re-exported. */
export type MeProfile = Pick<
  User,
  | 'id'
  | 'displayName'
  | 'mail'
  | 'userPrincipalName'
  | 'jobTitle'
  | 'department'
  | 'officeLocation'
>;

export class MeResource {
  constructor(private readonly graph: Client) {}

  /** GET /me — basic profile for the signed-in user. */
  async get(): Promise<MeProfile> {
    try {
      return (await this.graph.api('/me').get()) as MeProfile;
    } catch (err) {
      throw liftGraphError(err);
    }
  }
}
```

```ts
// packages/core/src/client.ts
import type { Client } from '@microsoft/microsoft-graph-client';

import { MeResource } from './resources/me.js';

/**
 * Top-level facade over the Graph endpoints we expose. New resources
 * (mail, calendar, contacts) hang off this client in subsequent slices.
 */
export class OutlookClient {
  public readonly me: MeResource;

  constructor(graph: Client) {
    this.me = new MeResource(graph);
  }
}
```

- [ ] **Step 4: Re-export from the barrel**

Append to `packages/core/src/index.ts`:

```ts
export { OutlookClient } from './client.js';
export { makeGraphClient } from './graph/client.js';
export type { MakeGraphClientOptions } from './graph/client.js';
export { MeResource } from './resources/me.js';
export type { MeProfile } from './resources/me.js';
export {
  CoreError,
  NotFoundError,
  ThrottledError,
  ServerError,
  NetworkError,
  liftGraphError,
} from './graph/errors.js';
```

- [ ] **Step 5: Run the tests + typecheck**

```bash
pnpm -F @alavida-ai/outlook-core vitest run src/client.test.ts && pnpm -F @alavida-ai/outlook-core typecheck
```

Expected: 2 tests pass, typecheck passes.

- [ ] **Step 6: Build the package once to make sure dist generates cleanly**

```bash
pnpm -F @alavida-ai/outlook-core build
ls packages/core/dist/index.js packages/core/dist/index.d.ts
```

Expected: both files present.

- [ ] **Step 7: Commit**

```bash
git add packages/core/src/resources/me.ts packages/core/src/client.ts packages/core/src/client.test.ts packages/core/src/index.ts
git commit -m "$(cat <<'EOF'
core: OutlookClient + MeResource (GET /me)

First wired resource. .me.get() calls the SDK and lifts any GraphError
into the local taxonomy. Subsequent slices add MailResource,
CalendarResource, ContactsResource the same way.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 3 — `cli` package: auth commands + whoami

### Task 3.1: `packages/cli` scaffold

**Files:**
- Create: `packages/cli/package.json`
- Create: `packages/cli/tsconfig.json`
- Create: `packages/cli/src/index.ts` (skeleton with help only)

- [ ] **Step 1: Write `packages/cli/package.json`**

```jsonc
{
  "name": "@alavida-ai/outlook-cli",
  "version": "0.1.0",
  "description": "Alavida Outlook CLI — read mail, draft messages, manage calendar via Microsoft Graph.",
  "type": "module",
  "license": "UNLICENSED",
  "bin": { "outlook": "dist/index.js" },
  "files": ["dist", "README.md"],
  "publishConfig": {
    "registry": "https://npm.pkg.github.com",
    "access": "restricted"
  },
  "scripts": {
    "build": "tsc -p tsconfig.json && chmod +x dist/index.js",
    "typecheck": "tsc -p tsconfig.json --noEmit"
  },
  "dependencies": {
    "@alavida-ai/outlook-core": "workspace:*"
  },
  "engines": { "node": ">=20" }
}
```

- [ ] **Step 2: Write `packages/cli/tsconfig.json`**

```jsonc
{
  "extends": "../../tsconfig.base.json",
  "compilerOptions": {
    "rootDir": "src",
    "outDir": "dist",
    "tsBuildInfoFile": "tsconfig.tsbuildinfo"
  },
  "include": ["src/**/*.ts"],
  "exclude": ["src/**/*.test.ts"],
  "references": [{ "path": "../core" }]
}
```

- [ ] **Step 3: Write skeleton `src/index.ts`**

```ts
#!/usr/bin/env node
/**
 * outlook — Microsoft Outlook CLI on top of @alavida-ai/outlook-core.
 *
 * Stdout = data. Stderr = human messages. Exit 0 success, 1 user/auth
 * error, 2 unexpected error.
 *
 * Uses Node's util.parseArgs — no CLI-framework dep.
 */
import { eprintln, formatError } from './output.js';

const TOP_HELP = `Usage: outlook <command> [args...]

Commands:
  auth login         Sign in via device-code flow.
  auth logout        Clear cached tokens.
  auth status        Show the signed-in account, if any.
  whoami             Print the signed-in user's profile.

Global flags:
  --account UPN      Pick a specific cached account (or set OUTLOOK_ACCOUNT).
  --json             Emit JSON to stdout instead of a human summary.

Environment:
  AZURE_CLIENT_ID    Override the embedded Entra app id.
  AZURE_TENANT_ID    Override the default tenant ('common').
  OUTLOOK_ACCOUNT    Default UPN to use when multiple accounts cached.

Run \`outlook <command> --help\` for command-specific options.
`;

async function main(argv: string[]): Promise<number> {
  const [first, second, ...rest] = argv;

  if (!first || first === '--help' || first === '-h' || first === 'help') {
    process.stdout.write(TOP_HELP);
    return first ? 0 : 1;
  }

  if (first === 'auth') {
    if (second === 'login') {
      const { run } = await import('./commands/auth-login.js');
      return run(rest);
    }
    if (second === 'logout') {
      const { run } = await import('./commands/auth-logout.js');
      return run(rest);
    }
    if (second === 'status') {
      const { run } = await import('./commands/auth-status.js');
      return run(rest);
    }
    eprintln(`Unknown auth subcommand: ${second ?? '(none)'}.`);
    return 1;
  }

  if (first === 'whoami') {
    const { run } = await import('./commands/whoami.js');
    return run([second, ...rest].filter((v): v is string => v !== undefined));
  }

  eprintln(`Unknown command: ${first}. Run \`outlook --help\`.`);
  return 1;
}

main(process.argv.slice(2))
  .then((code) => process.exit(code))
  .catch((err: unknown) => {
    eprintln(formatError(err));
    process.exit(2);
  });
```

- [ ] **Step 4: Install and typecheck (typecheck will fail until next tasks land — that's OK to verify the linkage)**

```bash
pnpm install
pnpm -F @alavida-ai/outlook-cli typecheck 2>&1 | head -20
```

Expected: errors referencing `./output.js` and `./commands/*` — modules-not-found. (We'll fill these in the next steps.)

- [ ] **Step 5: Stage everything; commit after Task 3.4 when the package typechecks clean**

```bash
git add packages/cli/package.json packages/cli/tsconfig.json packages/cli/src/index.ts pnpm-lock.yaml
```

---

### Task 3.2: `cli` output helpers + error formatter

**Files:**
- Create: `packages/cli/src/output.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/cli/src/output.ts
import {
  AuthError,
  CoreError,
  NetworkError,
  NotFoundError,
  ServerError,
  ThrottledError,
} from '@alavida-ai/outlook-core';

/** Write a JSON payload to stdout, terminated with a newline. */
export function printJson(value: unknown): void {
  process.stdout.write(JSON.stringify(value, null, 2) + '\n');
}

/** Write a line of text to stderr. */
export function eprintln(line = ''): void {
  process.stderr.write(line + '\n');
}

/** Write a line of text to stdout. */
export function println(line = ''): void {
  process.stdout.write(line + '\n');
}

/**
 * Format any thrown error for stderr. AuthError surfaces .nextStep
 * verbatim; CoreError variants get a one-line summary.
 */
export function formatError(e: unknown): string {
  if (e instanceof AuthError) {
    return `${e.message}\n  Next: ${e.nextStep}`;
  }
  if (e instanceof NotFoundError) {
    return `Not found: ${e.message}`;
  }
  if (e instanceof ThrottledError) {
    const ra = e.retryAfterSeconds !== null ? ` Retry after ${e.retryAfterSeconds}s.` : '';
    return `Microsoft Graph throttled the request.${ra}`;
  }
  if (e instanceof ServerError) {
    return `Microsoft Graph server error: ${e.message}`;
  }
  if (e instanceof NetworkError) {
    return `Network error: ${e.message}`;
  }
  if (e instanceof CoreError) {
    return `Outlook core error: ${e.message}`;
  }
  if (e instanceof Error) return `Unexpected error: ${e.message}`;
  return `Unexpected error: ${String(e)}`;
}
```

- [ ] **Step 2: Stage**

```bash
git add packages/cli/src/output.ts
```

---

### Task 3.3: Shared CLI client + flag parsing

**Files:**
- Create: `packages/cli/src/client.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/cli/src/client.ts
import { homedir } from 'node:os';
import { join } from 'node:path';
import {
  buildMsalApp,
  FileTokenCache,
  makeGraphClient,
  OutlookClient,
  type TokenCache,
} from '@alavida-ai/outlook-core';
import type { PublicClientApplication } from '@azure/msal-node';

export interface CliContext {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn: string | undefined;
  outlook: OutlookClient;
}

/** Resolve the preferred UPN: --account flag wins, then OUTLOOK_ACCOUNT env. */
export function resolveUpn(accountFlag: string | undefined): string | undefined {
  if (accountFlag) return accountFlag;
  const env = process.env.OUTLOOK_ACCOUNT;
  return env ? env : undefined;
}

/** Resolve the default token-cache path: $OUTLOOK_TOKEN_CACHE > ~/.outlook-cli/tokens.json */
export function defaultCachePath(): string {
  const override = process.env.OUTLOOK_TOKEN_CACHE;
  if (override) return override;
  return join(homedir(), '.outlook-cli', 'tokens.json');
}

export interface MakeContextOptions {
  preferredUpn?: string;
  cachePath?: string;
  clientId?: string;
  tenantId?: string;
}

/**
 * Construct an MSAL app + token cache + Graph client wired together.
 *
 * Every CLI command that hits Graph calls this. Cheap to construct — none
 * of the children touch disk or the network until the first method call.
 */
export function makeContext(opts: MakeContextOptions = {}): CliContext {
  const cache: TokenCache = new FileTokenCache(opts.cachePath ?? defaultCachePath());
  const app = buildMsalApp({
    cache,
    clientId: opts.clientId ?? process.env.AZURE_CLIENT_ID,
    tenantId: opts.tenantId ?? process.env.AZURE_TENANT_ID,
  });
  const preferredUpn = opts.preferredUpn;
  const graph = makeGraphClient({ app, cache, preferredUpn });
  return { app, cache, preferredUpn, outlook: new OutlookClient(graph) };
}
```

- [ ] **Step 2: Stage**

```bash
git add packages/cli/src/client.ts
```

---

### Task 3.4: `outlook auth login` command

**Files:**
- Create: `packages/cli/src/commands/auth-login.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/cli/src/commands/auth-login.ts
import { parseArgs } from 'node:util';
import { loginDeviceCode } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook auth login [--json]

Run the Microsoft device-code flow and cache the resulting tokens.

  - URL + code are printed to STDERR on the first line. Forward those
    to a human and wait for sign-in to complete.
  - Tokens are cached at ~/.outlook-cli/tokens.json (0600) — set
    OUTLOOK_TOKEN_CACHE to override.
  - With --json, the result is emitted to stdout as
    { "account": "<upn>" }.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
    });
  } catch (err) {
    eprintln(formatError(err));
    eprintln(HELP);
    return 1;
  }

  if (parsed.values.help) {
    process.stdout.write(HELP);
    return 0;
  }

  const ctx = makeContext();

  try {
    const result = await loginDeviceCode({
      app: ctx.app,
      cache: ctx.cache,
      onDeviceCode: (info) => {
        // First-line-of-stderr UX preserved: the full message includes URL + code.
        eprintln(info.message);
      },
    });
    if (parsed.values.json) {
      printJson({ account: result.account.username });
    } else {
      println(`Signed in as ${result.account.username}.`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
```

- [ ] **Step 2: Typecheck the cli package**

```bash
pnpm -F @alavida-ai/outlook-cli typecheck 2>&1 | head -20
```

Expected: still fails — `./commands/auth-logout.js` / `auth-status.js` / `whoami.js` missing. (Each is a separate task; we'll consolidate the commit at the end of Phase 3.)

- [ ] **Step 3: Stage**

```bash
git add packages/cli/src/commands/auth-login.ts
```

---

### Task 3.5: `outlook auth logout` command

**Files:**
- Create: `packages/cli/src/commands/auth-logout.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/cli/src/commands/auth-logout.ts
import { parseArgs } from 'node:util';
import { logout } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook auth logout [--json]

Remove the cached tokens and lock file. Idempotent.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
    });
  } catch (err) {
    eprintln(formatError(err));
    eprintln(HELP);
    return 1;
  }

  if (parsed.values.help) {
    process.stdout.write(HELP);
    return 0;
  }

  const ctx = makeContext();
  try {
    await logout({ cache: ctx.cache });
    if (parsed.values.json) {
      printJson({ status: 'logged_out' });
    } else {
      println('Logged out.');
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
```

- [ ] **Step 2: Stage**

```bash
git add packages/cli/src/commands/auth-logout.ts
```

---

### Task 3.6: `outlook auth status` command

**Files:**
- Create: `packages/cli/src/commands/auth-status.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/cli/src/commands/auth-status.ts
import { parseArgs } from 'node:util';
import { status } from '@alavida-ai/outlook-core';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook auth status [--account UPN] [--json]

Print the cached account, or exit 1 if not logged in.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        account: { type: 'string' },
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
    });
  } catch (err) {
    eprintln(formatError(err));
    eprintln(HELP);
    return 1;
  }

  if (parsed.values.help) {
    process.stdout.write(HELP);
    return 0;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const account = await status({ app: ctx.app, cache: ctx.cache, preferredUpn });
    if (!account) {
      if (parsed.values.json) {
        printJson({ logged_in: false });
      } else {
        eprintln('Not logged in. Run `outlook auth login`.');
      }
      return 1;
    }
    if (parsed.values.json) {
      printJson({
        logged_in: true,
        username: account.username,
        tenantId: account.tenantId,
        homeAccountId: account.homeAccountId,
      });
    } else {
      println(`Signed in as ${account.username} (tenant ${account.tenantId}).`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
```

- [ ] **Step 2: Stage**

```bash
git add packages/cli/src/commands/auth-status.ts
```

---

### Task 3.7: `outlook whoami` command

**Files:**
- Create: `packages/cli/src/commands/whoami.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/cli/src/commands/whoami.ts
import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook whoami [--account UPN] [--json]

Print the signed-in user's display name, email, job title, department,
and office location.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        account: { type: 'string' },
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
    });
  } catch (err) {
    eprintln(formatError(err));
    eprintln(HELP);
    return 1;
  }

  if (parsed.values.help) {
    process.stdout.write(HELP);
    return 0;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const me = await ctx.outlook.me.get();
    if (parsed.values.json) {
      printJson({
        id: me.id ?? null,
        displayName: me.displayName ?? null,
        mail: me.mail ?? null,
        userPrincipalName: me.userPrincipalName ?? null,
        jobTitle: me.jobTitle ?? null,
        department: me.department ?? null,
        officeLocation: me.officeLocation ?? null,
      });
    } else {
      println(me.displayName ?? '(no display name)');
      println(`  Email:      ${me.mail ?? me.userPrincipalName ?? '(none)'}`);
      println(`  Job title:  ${me.jobTitle ?? '-'}`);
      println(`  Department: ${me.department ?? '-'}`);
      println(`  Office:     ${me.officeLocation ?? '-'}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
```

- [ ] **Step 2: Typecheck the cli package — should now pass clean**

```bash
pnpm -F @alavida-ai/outlook-cli typecheck
```

Expected: no errors.

- [ ] **Step 3: Build it and verify the bin is executable**

```bash
pnpm -F @alavida-ai/outlook-cli build
ls -l packages/cli/dist/index.js
```

Expected: `dist/index.js` exists and is executable (`chmod +x` ran from the build script).

- [ ] **Step 4: Smoke-test `--help`**

```bash
node packages/cli/dist/index.js --help
```

Expected: the top-level help text on stdout, exit 0.

- [ ] **Step 5: Commit Phase 3 in one go**

```bash
git add packages/cli/src/commands/whoami.ts
git commit -m "$(cat <<'EOF'
cli: scaffold + auth login/logout/status + whoami

Per spec §Package responsibilities — cli row. Layout copied from
granola: util.parseArgs (no CLI-framework dep), src/commands/<verb>.ts
each exporting run(argv): Promise<number>, lazy await import() from
src/index.ts.

  - outlook auth login: device-code flow; URL+code on stderr first
    line; --json for { account } output
  - outlook auth logout: idempotent, removes tokens.json + lock file
  - outlook auth status: prints cached account; exit 1 if not logged in
  - outlook whoami: GET /me; pretty multi-line by default, --json for
    structured payload

--account / OUTLOOK_ACCOUNT resolution is in src/client.ts so every
command picks it up the same way.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 3.8: Manual smoke test of `outlook auth login` + `outlook whoami`

This is a sanity check before we move on. **You'll be asked to sign in interactively** — that's the device-code flow doing its job.

- [ ] **Step 1: Make sure no prior cache exists**

```bash
rm -f ~/.outlook-cli/tokens.json ~/.outlook-cli/tokens.json.lock
node packages/cli/dist/index.js auth status ; echo "exit=$?"
```

Expected: `Not logged in. Run \`outlook auth login\`.` on stderr, `exit=1`.

- [ ] **Step 2: Run the device-code flow**

```bash
node packages/cli/dist/index.js auth login
```

Expected: a single line on stderr like `To sign in, use a web browser to open https://microsoft.com/devicelogin and enter the code XXXXXXXX...`. Open the URL, enter the code, sign in with your Microsoft account. The command then prints `Signed in as <your-upn>.` on stdout and exits 0.

- [ ] **Step 3: Verify status**

```bash
node packages/cli/dist/index.js auth status
```

Expected: `Signed in as <your-upn> (tenant <tenant-id>).` on stdout, exit 0.

- [ ] **Step 4: Verify `whoami`**

```bash
node packages/cli/dist/index.js whoami
```

Expected: your display name, email, job title, department, office location — each on its own line.

- [ ] **Step 5: Verify `whoami --json`**

```bash
node packages/cli/dist/index.js whoami --json | jq .
```

Expected: a JSON object with the same fields. `jq` parses cleanly.

- [ ] **Step 6: Verify the token cache file permissions**

```bash
ls -la ~/.outlook-cli/
```

Expected: directory `0700`, `tokens.json` `0600`.

**Do NOT commit anything from this task** — it's verification only. If any step fails, stop and debug before moving on.

---

## Phase 4 — `openclaw` package: scaffold + `whoami` tool

### Task 4.1: `packages/openclaw` scaffold + plugin manifest

**Files:**
- Create: `packages/openclaw/package.json`
- Create: `packages/openclaw/tsconfig.json`
- Create: `packages/openclaw/openclaw.plugin.json`
- Create: `packages/openclaw/skills/outlook/.gitkeep` (placeholder; SKILL.md moves in the final slice)

- [ ] **Step 1: Write `packages/openclaw/package.json`**

```jsonc
{
  "name": "@alavida-ai/outlook-plugin-openclaw",
  "version": "0.1.0",
  "description": "OpenClaw plugin: native tools for reading mail, drafting messages, and managing calendar via Microsoft Graph.",
  "type": "module",
  "license": "UNLICENSED",
  "main": "dist/index.js",
  "homepage": "https://github.com/alavida-ai/outlook-cli",
  "repository": {
    "type": "git",
    "url": "https://github.com/alavida-ai/outlook-cli.git",
    "directory": "packages/openclaw"
  },
  "files": [
    "dist",
    "openclaw.plugin.json",
    "skills",
    "README.md"
  ],
  "publishConfig": {
    "registry": "https://npm.pkg.github.com",
    "access": "restricted"
  },
  "scripts": {
    "build": "tsc -p tsconfig.json",
    "typecheck": "tsc -p tsconfig.json --noEmit"
  },
  "dependencies": {
    "@alavida-ai/outlook-core": "workspace:*",
    "typebox": "^1.1.38"
  },
  "peerDependencies": {
    "openclaw": ">=2026.3.24-beta.2"
  },
  "devDependencies": {
    "openclaw": "2026.5.12"
  },
  "openclaw": {
    "extensions": ["./dist/index.js"],
    "compat": {
      "pluginApi": ">=2026.3.24-beta.2",
      "minGatewayVersion": "2026.3.24-beta.2"
    },
    "build": {
      "openclawVersion": "2026.3.24-beta.2",
      "pluginSdkVersion": "2026.3.24-beta.2"
    }
  },
  "engines": { "node": ">=20" }
}
```

- [ ] **Step 2: Write `packages/openclaw/tsconfig.json`**

```jsonc
{
  "extends": "../../tsconfig.base.json",
  "compilerOptions": {
    "rootDir": "src",
    "outDir": "dist",
    "tsBuildInfoFile": "tsconfig.tsbuildinfo"
  },
  "include": ["src/**/*.ts"],
  "exclude": ["src/**/*.test.ts"],
  "references": [{ "path": "../core" }]
}
```

- [ ] **Step 3: Write `packages/openclaw/openclaw.plugin.json`**

```json
{
  "id": "outlook",
  "name": "outlook",
  "description": "Read and triage Microsoft Outlook mail + calendar via Microsoft Graph as the signed-in user. Delegated permissions; draft-only mail.",
  "skills": ["./skills"],
  "contracts": {
    "tools": ["whoami"]
  },
  "activation": { "onStartup": true },
  "configSchema": {
    "type": "object",
    "additionalProperties": false,
    "properties": {
      "clientId": {
        "type": "string",
        "description": "Override the embedded Entra app id."
      },
      "tenantId": {
        "type": "string",
        "description": "Override the default tenant ('common')."
      },
      "tokenCachePath": {
        "type": "string",
        "description": "Path to the MSAL token cache. Defaults to ~/.outlook-cli/tokens.json."
      },
      "account": {
        "type": "string",
        "description": "UPN to use when multiple accounts are cached. Required for multi-account hosts."
      }
    }
  }
}
```

- [ ] **Step 4: Create the skills placeholder**

```bash
mkdir -p packages/openclaw/skills/outlook
touch packages/openclaw/skills/outlook/.gitkeep
```

(The full `SKILL.md` move happens in the last slice, after the tool surface is complete.)

- [ ] **Step 5: Install (adds typebox + openclaw devDep)**

```bash
pnpm install
```

Expected: lockfile updated. Watch for any `allowBuilds` surprises in the output — if pnpm suggests new install scripts (e.g. `koffi`), accept them in `pnpm-workspace.yaml` first; granola already enumerates the ones we expect.

- [ ] **Step 6: Stage but don't commit yet**

```bash
git add packages/openclaw/package.json packages/openclaw/tsconfig.json packages/openclaw/openclaw.plugin.json packages/openclaw/skills/outlook/.gitkeep pnpm-lock.yaml
```

---

### Task 4.2: Port `register.ts`, `shared-schemas.ts`, `errors.ts`, `pretty.ts` from granola

These four files are deliberately copy-paste-with-rename from `granola-plugin/packages/openclaw/src/` — they're the cross-cutting plumbing every Alavida OpenClaw plugin will share. No new logic.

**Files:**
- Create: `packages/openclaw/src/shared-schemas.ts`
- Create: `packages/openclaw/src/errors.ts`
- Create: `packages/openclaw/src/pretty.ts`
- Create: `packages/openclaw/src/register.ts`

- [ ] **Step 1: Copy `shared-schemas.ts` from granola verbatim**

```bash
cp /Users/alexandergarciachicote/code/projects/granola-plugin/packages/openclaw/src/shared-schemas.ts packages/openclaw/src/shared-schemas.ts
```

- [ ] **Step 2: Copy `errors.ts` from granola, rename `GranolaError` → no rename required (the file imports `@alavida-ai/granola-core`; we'll edit one import)**

```bash
cp /Users/alexandergarciachicote/code/projects/granola-plugin/packages/openclaw/src/errors.ts packages/openclaw/src/errors.ts
```

Then open `packages/openclaw/src/errors.ts` and replace the granola error imports with our error types. Specifically:

- Change `from '@alavida-ai/granola-core'` to `from '@alavida-ai/outlook-core'`
- Replace any `GranolaAuthError` / `GranolaNotFoundError` / `GranolaRateLimitError` / `GranolaServerError` / `GranolaNetworkError` references with our equivalents: `AuthError`, `NotFoundError`, `ThrottledError`, `ServerError`, `NetworkError` respectively.
- Replace any `GranolaError` (the parent) references with `CoreError`.

Open the granola file first to see the exact structure before editing; the changes are mechanical.

- [ ] **Step 3: Copy `pretty.ts` from granola**

```bash
cp /Users/alexandergarciachicote/code/projects/granola-plugin/packages/openclaw/src/pretty.ts packages/openclaw/src/pretty.ts
```

The granola version is essentially a switch-on-shape dispatcher. Leave it as-is for now; we add an Outlook-specific case in Task 4.5 (or just rely on the generic JSON fallback for `whoami`).

- [ ] **Step 4: Copy `register.ts` from granola**

```bash
cp /Users/alexandergarciachicote/code/projects/granola-plugin/packages/openclaw/src/register.ts packages/openclaw/src/register.ts
```

Then in `packages/openclaw/src/register.ts`, change the `PluginConfig` import from `from './client.js'` to keep that path (we'll create our own `client.ts` next task with a typebox-validated `PluginConfig`) and adapt `readPluginConfig` to read our four config fields:

Replace the body of `readPluginConfig` with:

```ts
export function readPluginConfig(api: OpenClawPluginApi): PluginConfig {
  const raw = (api.pluginConfig ?? {}) as Record<string, unknown>;
  return {
    clientId: typeof raw.clientId === 'string' ? raw.clientId : undefined,
    tenantId: typeof raw.tenantId === 'string' ? raw.tenantId : undefined,
    tokenCachePath: typeof raw.tokenCachePath === 'string' ? raw.tokenCachePath : undefined,
    account: typeof raw.account === 'string' ? raw.account : undefined,
  };
}
```

- [ ] **Step 5: Stage**

```bash
git add packages/openclaw/src/shared-schemas.ts packages/openclaw/src/errors.ts packages/openclaw/src/pretty.ts packages/openclaw/src/register.ts
```

---

### Task 4.3: `openclaw` package `PluginConfig` + `getClient`

**Files:**
- Create: `packages/openclaw/src/client.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/openclaw/src/client.ts
import { homedir } from 'node:os';
import { join } from 'node:path';
import {
  buildMsalApp,
  FileTokenCache,
  makeGraphClient,
  OutlookClient,
  type TokenCache,
} from '@alavida-ai/outlook-core';

export interface PluginConfig {
  clientId?: string;
  tenantId?: string;
  tokenCachePath?: string;
  account?: string;
}

function defaultCachePath(): string {
  return process.env.OUTLOOK_TOKEN_CACHE ?? join(homedir(), '.outlook-cli', 'tokens.json');
}

let cachedClient: { config: PluginConfig; client: OutlookClient } | null = null;

/**
 * Build (or reuse) the OutlookClient for the supplied plugin config.
 *
 * The plugin SDK calls our tools many times per session; constructing the
 * MSAL app and FileTokenCache on every call is wasteful. We memoise on a
 * structural-equality check of the config; if the operator hot-reloads the
 * config with new values, the next call rebuilds.
 */
export function getClient(config: PluginConfig): OutlookClient {
  if (cachedClient && shallowEqualConfig(cachedClient.config, config)) {
    return cachedClient.client;
  }
  const cache: TokenCache = new FileTokenCache(config.tokenCachePath ?? defaultCachePath());
  const app = buildMsalApp({
    cache,
    clientId: config.clientId,
    tenantId: config.tenantId,
  });
  const graph = makeGraphClient({ app, cache, preferredUpn: config.account });
  const client = new OutlookClient(graph);
  cachedClient = { config: { ...config }, client };
  return client;
}

/** Test-only: reset the memoised client. Exported so unit tests can isolate. */
export function _resetClientForTesting(): void {
  cachedClient = null;
}

function shallowEqualConfig(a: PluginConfig, b: PluginConfig): boolean {
  return (
    a.clientId === b.clientId &&
    a.tenantId === b.tenantId &&
    a.tokenCachePath === b.tokenCachePath &&
    a.account === b.account
  );
}
```

- [ ] **Step 2: Stage**

```bash
git add packages/openclaw/src/client.ts
```

---

### Task 4.4: `whoami` tool

**Files:**
- Create: `packages/openclaw/src/tools/whoami.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/openclaw/src/tools/whoami.ts
/**
 * `whoami` — display the signed-in user's profile.
 *
 * Mirrors `outlook whoami` on the CLI. Pulls /me from Graph and returns a
 * compact subset (id, displayName, mail, userPrincipalName, jobTitle,
 * department, officeLocation). Read-only.
 */
import { Type } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

interface WhoamiResult {
  id: string | null;
  displayName: string | null;
  mail: string | null;
  userPrincipalName: string | null;
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
}

const whoami = defineTool({
  name: 'whoami',
  description:
    "Display the signed-in user's basic profile (display name, email, job title, department, office). Read-only.",
  parameters: Type.Object({}),
  async execute(_params, config): Promise<WhoamiResult> {
    const client = getClient(config);
    const me = await client.me.get();
    return {
      id: me.id ?? null,
      displayName: me.displayName ?? null,
      mail: me.mail ?? null,
      userPrincipalName: me.userPrincipalName ?? null,
      jobTitle: me.jobTitle ?? null,
      department: me.department ?? null,
      officeLocation: me.officeLocation ?? null,
    };
  },
});

export default whoami;
```

- [ ] **Step 2: Stage**

```bash
git add packages/openclaw/src/tools/whoami.ts
```

---

### Task 4.5: Plugin entry — register tools, plumb config

**Files:**
- Create: `packages/openclaw/src/index.ts`

- [ ] **Step 1: Write the implementation**

```ts
// packages/openclaw/src/index.ts
/**
 * OpenClaw plugin entry — outlook.
 *
 * Reads and triages Outlook mail + calendar via Microsoft Graph as the
 * signed-in user. Delegated permissions; draft-only mail.
 *
 * Architecture:
 *   - One file per tool in `./tools/<tool-name>.ts`, each default-exporting
 *     a {@link ToolDescriptor}.
 *   - {@link registerTool} wraps every descriptor with shared output/help
 *     injection, pretty/json dispatch, and `withErrorMapping`.
 *
 * This first slice ships `whoami` only. Mail, calendar, and contacts tools
 * land in subsequent plan files.
 */
import { Type } from 'typebox';
import { definePluginEntry } from 'openclaw/plugin-sdk/plugin-entry';

import type { PluginConfig } from './client.js';
import { readPluginConfig, registerTool, type ToolDescriptor } from './register.js';

import whoami from './tools/whoami.js';

const TOOLS: ToolDescriptor[] = [whoami];

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const configJsonSchema: any = Type.Object({
  clientId: Type.Optional(
    Type.String({ description: 'Override the embedded Entra app id.' }),
  ),
  tenantId: Type.Optional(
    Type.String({ description: "Override the default tenant ('common')." }),
  ),
  tokenCachePath: Type.Optional(
    Type.String({
      description: 'Path to the MSAL token cache. Defaults to ~/.outlook-cli/tokens.json.',
    }),
  ),
  account: Type.Optional(
    Type.String({
      description: 'UPN to use when multiple accounts are cached. Required for multi-account hosts.',
    }),
  ),
});

export default definePluginEntry({
  id: 'outlook',
  name: 'outlook',
  description:
    'Read and triage Outlook mail + calendar via Microsoft Graph as the signed-in user. Delegated permissions; draft-only mail.',
  configSchema: { jsonSchema: configJsonSchema },
  register(api) {
    const getConfig = () => readPluginConfig(api);
    for (const tool of TOOLS) {
      registerTool(api, tool, getConfig);
    }
  },
});

export { getClient, _resetClientForTesting } from './client.js';
export type { PluginConfig };
export { withErrorMapping, isToolErrorEnvelope } from './errors.js';
export type { ToolErrorEnvelope, ToolErrorResponse } from './errors.js';
export { registerTool, defineTool, type ToolDescriptor } from './register.js';
```

- [ ] **Step 2: Typecheck the openclaw package**

```bash
pnpm -F @alavida-ai/outlook-plugin-openclaw typecheck
```

Expected: passes. If anything fails because `pretty.ts` imports a granola-only shape, replace those imports with no-ops; pretty rendering is optional for v1 (the JSON fallback covers it).

- [ ] **Step 3: Build the package**

```bash
pnpm -F @alavida-ai/outlook-plugin-openclaw build
ls packages/openclaw/dist/index.js
```

Expected: file exists.

- [ ] **Step 4: Run the full workspace tests + typecheck for regression**

```bash
pnpm typecheck && pnpm test
```

Expected: clean across all three packages.

- [ ] **Step 5: Commit Phase 4**

```bash
git add packages/openclaw/src/index.ts
git commit -m "$(cat <<'EOF'
openclaw: scaffold plugin + whoami tool

Per spec §Package responsibilities — openclaw row, and
openclaw-plugin-distribution.md for the manifest shape.

Plugin entry registers a single tool for this slice (`whoami`), wired to
@alavida-ai/outlook-core via getClient(). PluginConfig accepts clientId
/ tenantId / tokenCachePath / account overrides per spec §4.

Cross-cutting plumbing (register.ts, shared-schemas.ts, errors.ts,
pretty.ts) ported from granola-plugin with imports retargeted to
@alavida-ai/outlook-core. Memoised getClient avoids rebuilding MSAL
on every tool call.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 4.6: Smoke-test the OpenClaw plugin locally

This validates that the plugin loads in a real OpenClaw gateway and the `whoami` tool calls Graph successfully.

- [ ] **Step 1: Pack the plugin into a local tarball**

```bash
pnpm -F @alavida-ai/outlook-plugin-openclaw build
cd packages/openclaw && pnpm pack && cd -
ls packages/openclaw/*.tgz
```

Expected: a `.tgz` file like `packages/openclaw/alavida-ai-outlook-plugin-openclaw-0.1.0.tgz`.

- [ ] **Step 2: Install the local tarball into your OpenClaw host**

If you have OpenClaw running locally:

```bash
openclaw plugins install ./packages/openclaw/alavida-ai-outlook-plugin-openclaw-0.1.0.tgz
openclaw gateway restart
openclaw plugins inspect outlook
```

Expected: plugin appears with one tool (`whoami`); no errors.

If you don't have OpenClaw running locally, **skip this step**. The CI gating for the plugin happens via `pnpm typecheck` + `pnpm build`; full integration testing waits until the gateway dry-run before the rewrite branch merges.

- [ ] **Step 3: Call the tool through OpenClaw**

```bash
openclaw tools call outlook whoami
```

Expected: a JSON payload with your `displayName`, `mail`, etc.

- [ ] **Step 4: Clean up the tarball — it's not committed**

```bash
rm packages/openclaw/*.tgz
```

**Do NOT commit anything from this task** — it's verification only.

---

## Closing

After Phase 4 the repo can do this:

- `outlook auth login` / `auth logout` / `auth status` end-to-end
- `outlook whoami` end-to-end
- `@alavida-ai/outlook-plugin-openclaw` loads in a real OpenClaw gateway and exposes a working `whoami` tool

The next plan file in this series (`2026-XX-XX-outlook-mail-read-slice.md`) adds the mail-read tools (`mail_list`, `mail_read`, `mail_search`, `mail_folders`, `mail_list_attachments`, `mail_download_attachment`) and their CLI subcommands following the same per-slice pattern: write the `core` resource method + tests → add the CLI subcommand → add the OpenClaw tool, each as its own commit, all in one PR.

Subsequent slices add: mail write/triage, calendar, contacts stub, then the final polish slice that moves `SKILL.md` into `packages/openclaw/skills/outlook/` and rewrites the README around the new layout.

---

## Self-Review

**Spec coverage:**
- §Architecture (3-package layout, granola precedent): Tasks 0.4–0.10, 1.1, 3.1, 4.1 ✅
- §Package responsibilities — core: Phases 1 + 2 ✅
- §Package responsibilities — cli: Phase 3 ✅
- §Package responsibilities — openclaw: Phase 4 ✅
- §Auth strategy (file-only): Tasks 1.4, 3.3, 4.3 ✅
- §4.1 Why file-only: discussed in spec; no code task needed ✅
- §4.2.1 Atomic writes: Task 1.4 ✅
- §4.2.2 O_EXCL lock: Task 1.5 ✅
- §4.2.3 Integrity check on read: Task 1.4 ✅
- §4.2.4 Error taxonomy: Task 1.2 ✅
- §4.2.5 Multi-account handling: Task 1.7 + threaded through 3.6, 3.7, 4.3 ✅
- §4.2.6 File permissions: Task 1.4 ✅
- §4.2.7 Idempotent logout: Task 1.10 ✅
- §4.2.8 Same Entra app id: Task 1.6 (`EMBEDDED_CLIENT_ID`) ✅
- §4.2.9 Scope set: Task 1.6 (`OUTLOOK_SCOPES`) ✅
- §4.2.10 Auth test harness: covered piecewise in Tasks 1.6, 1.8, 1.9; an end-to-end concurrency test ought to land but is genuinely heavyweight (needs spawning two subprocesses) — deferred to the mail-read slice when we have a real Graph endpoint to exercise. **Tracked here as gap.** ⚠️
- §4.3 TokenCache abstraction: Task 1.3 ✅
- §4.4 Threat model: documentation only, no code task ✅
- §OpenClaw tool surface — `whoami`: Task 4.4 (other tools deferred to slice plans) ✅
- §Tech stack — every row touched in Phase 0–4 ✅
- §Supply-chain hardening: Tasks 0.2, 0.4, 0.5, 0.10 ✅
- §Cutover plan commits 1–3: Tasks 0.1, 0.11, end of Phase 2 ✅
- §Cutover plan commits 4+ (vertical slices): this plan IS slice 1 (whoami); future plans cover the rest

**Placeholder scan:** Searched the plan for "TBD", "TODO", "implement later", "Similar to Task N", "appropriate error handling" — none found. The one explicit gap (cross-process concurrency end-to-end test) is named with a ⚠️ above rather than a vague placeholder.

**Type consistency:**
- `TokenCache` interface (Task 1.3) used identically by `FileTokenCache` (1.4), `getAccessToken` (1.8), `loginDeviceCode` (1.9), `logout` (1.10), `status` (1.10), `MsalAuthenticationProvider` (2.1), `makeGraphClient` (2.3), `makeContext` (3.3), `getClient` (4.3) — verified.
- `AuthError` family naming consistent across 1.2 → exports in 1.11 → `formatError` in 3.2 → `liftGraphError` in 2.2.
- `OutlookClient` constructor takes a `Client` (graph SDK type) — used identically by `makeContext` (3.3) and `getClient` (4.3).
- `PluginConfig` shape `{clientId?, tenantId?, tokenCachePath?, account?}` matches `openclaw.plugin.json` configSchema (Task 4.1), `readPluginConfig` (4.2), `getClient` (4.3), the index entry's typebox schema (4.5).
- `preferredUpn: string | undefined` (vs `?: string`) used consistently in 1.8, 1.9, 2.1, 2.3, 3.3.
