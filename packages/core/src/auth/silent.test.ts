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
