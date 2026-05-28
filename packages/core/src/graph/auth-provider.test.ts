import { describe, expect, it, vi } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from '../auth/cache.js';
import { MsalAuthenticationProvider } from './auth-provider.js';

describe('MsalAuthenticationProvider', () => {
  it('returns the access token from getAccessToken()', async () => {
    const cache = new InMemoryTokenCache();
    const app = {
      getTokenCache: () => ({
        getAllAccounts: async () => [
          { homeAccountId: 'h', environment: 'e', tenantId: 't', username: 'a@x.com', localAccountId: 'l' },
        ],
      }),
      acquireTokenSilent: vi.fn(async () => ({
        accessToken: 'graph-tok',
        expiresOn: new Date(Date.now() + 3_600_000),
      })),
    } as unknown as PublicClientApplication;

    const provider = new MsalAuthenticationProvider({ app, cache, preferredUpn: undefined });
    expect(await provider.getAccessToken()).toBe('graph-tok');
  });
});
