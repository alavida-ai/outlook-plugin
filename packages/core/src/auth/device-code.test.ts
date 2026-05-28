import { describe, expect, it, vi } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from './cache.js';
import { loginDeviceCode } from './device-code.js';

describe('loginDeviceCode', () => {
  it('forwards the device-code message via the callback and resolves with the result', async () => {
    const cache = new InMemoryTokenCache();
    const messages: string[] = [];
    const acquireMock = vi.fn(async (req: any) => {
      req.deviceCodeCallback({
        message: 'To sign in, use a web browser to open https://microsoft.com/devicelogin and enter the code AB12CD34',
        userCode: 'AB12CD34',
        verificationUri: 'https://microsoft.com/devicelogin',
        deviceCode: 'd',
        expiresIn: 900,
        interval: 5,
      });
      return {
        accessToken: 'tok',
        account: {
          homeAccountId: 'h',
          environment: 'e',
          tenantId: 't',
          username: 'a@x.com',
          localAccountId: 'l',
        },
        expiresOn: new Date(Date.now() + 3_600_000),
      };
    });
    const app = { acquireTokenByDeviceCode: acquireMock } as unknown as PublicClientApplication;

    const result = await loginDeviceCode({
      app,
      cache,
      onDeviceCode: (msg) => messages.push(msg.message),
    });

    expect(result.account.username).toBe('a@x.com');
    expect(messages).toHaveLength(1);
    expect(messages[0]).toContain('AB12CD34');
  });

  it('runs under cache.lock()', async () => {
    const cache = new InMemoryTokenCache();
    const lockSpy = vi.spyOn(cache, 'lock');
    const app = {
      acquireTokenByDeviceCode: vi.fn(async () => ({
        accessToken: 'tok',
        account: {
          homeAccountId: 'h',
          environment: 'e',
          tenantId: 't',
          username: 'a@x.com',
          localAccountId: 'l',
        },
        expiresOn: new Date(Date.now() + 1_000),
      })),
    } as unknown as PublicClientApplication;
    await loginDeviceCode({ app, cache, onDeviceCode: () => undefined });
    expect(lockSpy).toHaveBeenCalledTimes(1);
  });
});
