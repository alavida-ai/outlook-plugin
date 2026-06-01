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

  it('regression: lock() succeeds when the parent directory does not exist yet', async () => {
    // First-ever auth login on a fresh host: ~/.outlook-plugin/ doesn't exist.
    // loginDeviceCode calls cache.lock() before any save() has created the dir.
    // lock() must mkdir the parent before O_EXCL-creating the lock file,
    // otherwise it fails with ENOENT.
    const freshCache = new FileTokenCache(join(dir, 'nested', 'deeper', 'tokens.json'));
    const result = await freshCache.lock(async () => 'ok');
    expect(result).toBe('ok');
    const { existsSync } = await import('node:fs');
    expect(existsSync(join(dir, 'nested', 'deeper'))).toBe(true);
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
    // Drop a stale lock file at the cache's actual lockPath
    // (`<path>.lock` — for tokens.json that's tokens.json.lock, NOT tokens.lock).
    const lockPath = cache.lockPath;
    await writeFile(lockPath, '0');
    const { utimesSync } = await import('node:fs');
    const old = Date.now() / 1000 - 120; // 2 minutes ago
    utimesSync(lockPath, old, old);
    await expect(cache.lock(async () => 'taken', { maxLockAgeMs: 60_000 })).resolves.toBe(
      'taken',
    );
  });

  it('regression: stale-lock-detection actually runs (covers branch missed in original test)', async () => {
    // Tighter test of the same property: assert the lock file is REMOVED
    // by the stale-lock force-take path. If the stale-lock branch ever
    // stops running, this fails because the original stale file remains.
    const { utimesSync, existsSync, readFileSync } = await import('node:fs');
    const lockPath = cache.lockPath;
    await writeFile(lockPath, 'stale-pid');
    const old = Date.now() / 1000 - 120;
    utimesSync(lockPath, old, old);
    expect(existsSync(lockPath)).toBe(true);
    await cache.lock(async () => {
      // While we hold the lock, the file should contain OUR pid, not 'stale-pid'.
      expect(readFileSync(lockPath, 'utf8')).toBe(String(process.pid));
    }, { maxLockAgeMs: 60_000 });
    expect(existsSync(lockPath)).toBe(false);
  });
});
