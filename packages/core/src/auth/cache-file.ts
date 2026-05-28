import { randomBytes } from 'node:crypto';
import {
  chmod,
  mkdir,
  open,
  readFile,
  rename,
  rm,
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
