import { randomBytes } from 'node:crypto';
import { closeSync, mkdirSync, openSync } from 'node:fs';
import {
  chmod,
  mkdir,
  open,
  readFile,
  rename,
  rm,
  stat,
  unlink,
  writeFile,
} from 'node:fs/promises';
import { dirname } from 'node:path';
import { setTimeout as sleep } from 'node:timers/promises';

import { AuthCacheCorruptError, AuthLockTimeoutError } from './errors.js';
import type { TokenCache } from './cache.js';

const REQUIRED_KEYS = ['AccessToken', 'RefreshToken', 'IdToken', 'Account', 'AppMetadata'] as const;

const DEFAULT_TIMEOUT_MS = 30_000;
const DEFAULT_MAX_LOCK_AGE_MS = 60_000;

export interface LockOptions {
  /** Total time to keep retrying lock acquisition. Default 30s. */
  timeoutMs?: number;
  /**
   * If an existing lock file is older than this, presume the holder is dead
   * and force-take it. Default 60s.
   */
  maxLockAgeMs?: number;
}

/**
 * Default `TokenCache` backend: a single JSON file at `path`, written
 * atomically (`tmpfile → fsync → rename`) with 0600 permissions and a
 * 0700 parent directory.
 *
 * Cross-process serialisation via `O_EXCL` on `<path>.lock`. Exponential
 * backoff to a 30s timeout; stale-lock detection at 60s.
 */
export class FileTokenCache implements TokenCache {
  constructor(public readonly path: string) {}

  get lockPath(): string {
    return `${this.path}.lock`;
  }

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

  async lock<T>(fn: () => Promise<T>, opts: LockOptions = {}): Promise<T> {
    const timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    const maxLockAgeMs = opts.maxLockAgeMs ?? DEFAULT_MAX_LOCK_AGE_MS;
    const start = Date.now();
    let attempt = 0;

    // Ensure parent dir exists before O_EXCL create. Otherwise the first-ever
    // lock attempt on a fresh install fails with ENOENT (the parent dir is
    // normally created by save(), but a sign-in acquires the lock BEFORE the
    // cache plugin's afterCacheAccess fires).
    //
    // Synchronous mkdir so the function doesn't yield to the event loop
    // before the openSync — keeps the "first caller wins openSync" property
    // that concurrent callers rely on.
    const dir = dirname(this.path);
    mkdirSync(dir, { recursive: true, mode: 0o700 });

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
}

function isENOENT(err: unknown): boolean {
  return typeof err === 'object' && err !== null && (err as { code?: string }).code === 'ENOENT';
}

function isEEXIST(err: unknown): boolean {
  return typeof err === 'object' && err !== null && (err as { code?: string }).code === 'EEXIST';
}

function isMsalShape(value: unknown): boolean {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;
  return REQUIRED_KEYS.every((k) => k in v);
}
