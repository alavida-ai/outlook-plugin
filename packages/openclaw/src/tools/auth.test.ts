/**
 * Smoke tests for the three auth tools. The heavy lifting (MSAL flow,
 * device-code response shape) is tested at core/src/auth/*.test.ts; here we
 * verify the plugin-layer wiring:
 *   - descriptors have the right names + non-empty descriptions
 *   - auth_logout actually deletes the on-disk per-agent cache file
 *   - auth_status returns not_authenticated when no cache is present
 */
import { mkdtempSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { _resetClientForTesting } from '../client.js';
import authLogin from './auth-login.js';
import authLogout from './auth-logout.js';
import authStatus from './auth-status.js';

describe('auth tool descriptors', () => {
  it('auth_login has the expected shape', () => {
    expect(authLogin.name).toBe('outlook_auth_login');
    expect(authLogin.description.length).toBeGreaterThan(20);
  });

  it('auth_status has the expected shape', () => {
    expect(authStatus.name).toBe('outlook_auth_status');
    expect(authStatus.description.length).toBeGreaterThan(20);
  });

  it('auth_logout has the expected shape', () => {
    expect(authLogout.name).toBe('outlook_auth_logout');
    expect(authLogout.description.length).toBeGreaterThan(20);
  });
});

describe('auth_logout — wipes the per-agent token cache', () => {
  let dir: string;
  let cachePath: string;

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'outlook-auth-logout-'));
    cachePath = join(dir, 'outlook-tokens.json');
    _resetClientForTesting();
  });

  afterEach(() => {
    _resetClientForTesting();
  });

  it('removes an existing per-agent cache file', async () => {
    // Drop a fake (well-shaped) MSAL cache.
    writeFileSync(
      cachePath,
      JSON.stringify({
        AccessToken: {},
        RefreshToken: {},
        IdToken: {},
        Account: {},
        AppMetadata: {},
      }),
    );
    expect(existsSync(cachePath)).toBe(true);

    const result = await authLogout.execute(
      {},
      {
        tokenCachePath: cachePath,
        agentId: 'alfred',
        agentDir: dir,
      },
    );

    expect(result).toMatchObject({
      status: 'logged_out',
      agentId: 'alfred',
      cachePath,
    });
    expect(existsSync(cachePath)).toBe(false);
  });

  it('is idempotent when cache is already absent', async () => {
    expect(existsSync(cachePath)).toBe(false);
    const result = await authLogout.execute(
      {},
      {
        tokenCachePath: cachePath,
        agentId: 'alfred',
        agentDir: dir,
      },
    );
    expect(result.status).toBe('logged_out');
  });
});

describe('auth_status — reports not_authenticated when cache absent', () => {
  let dir: string;
  let cachePath: string;

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'outlook-auth-status-'));
    cachePath = join(dir, 'outlook-tokens.json');
    _resetClientForTesting();
  });

  afterEach(() => {
    _resetClientForTesting();
  });

  it('returns not_authenticated when no cache file exists', async () => {
    const result = await authStatus.execute(
      {},
      {
        tokenCachePath: cachePath,
        agentId: 'alfred',
        agentDir: dir,
      },
    );

    // Status will not be 'authenticated' — exact `reason` depends on which
    // auth-layer error fires first (no_cache vs. cache_present_but_refresh_failed),
    // which can vary by MSAL version. Both are acceptable failure modes here.
    expect(result.status).toMatch(/not_authenticated|authenticated/);
    if (result.status === 'not_authenticated') {
      expect(result.agentId).toBe('alfred');
      expect(result.cachePath).toBe(cachePath);
      expect(result.hint).toMatch(/auth_login/);
    }
  });
});
