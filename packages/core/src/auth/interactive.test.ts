import { describe, expect, it, vi } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from './cache.js';
import { AuthRefreshFailedError } from './errors.js';
import { loginInteractive } from './interactive.js';

const FAKE_ACCOUNT = {
  homeAccountId: 'h',
  environment: 'e',
  tenantId: 't',
  username: 'a@x.com',
  localAccountId: 'l',
};

describe('loginInteractive', () => {
  it('acquires a token interactively and returns the account + expiry', async () => {
    const cache = new InMemoryTokenCache();
    const acquireTokenInteractive = vi.fn(async () => ({
      accessToken: 'tok',
      account: FAKE_ACCOUNT,
      expiresOn: new Date(Date.now() + 3_600_000),
    }));
    const app = { acquireTokenInteractive } as unknown as PublicClientApplication;
    const openBrowser = vi.fn(async () => {});

    const result = await loginInteractive({ app, cache, openBrowser });

    expect(result.account.username).toBe('a@x.com');
    expect(result.expiresOn).toBeInstanceOf(Date);
    const req = acquireTokenInteractive.mock.calls[0][0] as Record<string, unknown>;
    expect(req.openBrowser).toBe(openBrowser);
    expect(req.scopes).toEqual([
      'Mail.ReadWrite',
      'Calendars.Read',
      'Calendars.Read.Shared',
      'User.Read',
    ]);
  });

  it('runs under cache.lock()', async () => {
    const cache = new InMemoryTokenCache();
    const lockSpy = vi.spyOn(cache, 'lock');
    const app = {
      acquireTokenInteractive: vi.fn(async () => ({
        accessToken: 'tok',
        account: FAKE_ACCOUNT,
        expiresOn: new Date(Date.now() + 3_600_000),
      })),
    } as unknown as PublicClientApplication;

    await loginInteractive({ app, cache, openBrowser: async () => {} });

    expect(lockSpy).toHaveBeenCalledTimes(1);
  });

  it('throws AuthRefreshFailedError when MSAL returns no account', async () => {
    const cache = new InMemoryTokenCache();
    const app = {
      acquireTokenInteractive: vi.fn(async () => ({
        accessToken: 'tok',
        account: null,
        expiresOn: null,
      })),
    } as unknown as PublicClientApplication;

    await expect(
      loginInteractive({ app, cache, openBrowser: async () => {} }),
    ).rejects.toBeInstanceOf(AuthRefreshFailedError);
  });
});
