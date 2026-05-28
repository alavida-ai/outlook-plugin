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
