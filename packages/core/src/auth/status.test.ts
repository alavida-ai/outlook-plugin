import { describe, expect, it } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from './cache.js';
import { status } from './status.js';
import { AuthAmbiguousAccountError } from './errors.js';

function appWith(usernames: string[]): PublicClientApplication {
  return {
    getTokenCache: () => ({
      getAllAccounts: async () =>
        usernames.map((u) => ({
          homeAccountId: `${u}-h`,
          environment: 'e',
          tenantId: 't',
          username: u,
          localAccountId: 'l',
        })),
    }),
  } as unknown as PublicClientApplication;
}

describe('status', () => {
  it('returns null when no accounts are cached', async () => {
    expect(await status({ app: appWith([]), cache: new InMemoryTokenCache() })).toBe(null);
  });

  it('returns the single cached account when there is one', async () => {
    const result = await status({ app: appWith(['a@x.com']), cache: new InMemoryTokenCache() });
    expect(result?.username).toBe('a@x.com');
  });

  it('throws AuthAmbiguousAccountError when multiple accounts and no preference', async () => {
    await expect(
      status({ app: appWith(['a@x.com', 'b@y.com']), cache: new InMemoryTokenCache() }),
    ).rejects.toBeInstanceOf(AuthAmbiguousAccountError);
  });

  it('honours preferredUpn when provided', async () => {
    const result = await status({
      app: appWith(['a@x.com', 'b@y.com']),
      cache: new InMemoryTokenCache(),
      preferredUpn: 'b@y.com',
    });
    expect(result?.username).toBe('b@y.com');
  });
});
