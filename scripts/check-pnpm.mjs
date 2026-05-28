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
