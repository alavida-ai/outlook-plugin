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
