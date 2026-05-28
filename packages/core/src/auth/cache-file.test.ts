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
