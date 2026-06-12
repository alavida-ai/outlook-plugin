import { describe, expect, it, vi } from 'vitest';
import type { DeviceCodeRequest, PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from './cache.js';
import { loginDeviceCodeInBackground } from './device-code-background.js';

const FAKE_ACCOUNT = {
  homeAccountId: 'h',
  environment: 'e',
  tenantId: 't',
  username: 'a@x.com',
  localAccountId: 'l',
};

describe('loginDeviceCodeInBackground', () => {
  it('resolves with verificationUrl + userCode as soon as MSAL emits the device-code response', async () => {
    const cache = new InMemoryTokenCache();
    // The MSAL fake: emit device code immediately, but stay "polling" until we release it.
    let release!: () => void;
    const pollingPromise = new Promise<void>((r) => {
      release = r;
    });
    const acquireMock = vi.fn(async (req: DeviceCodeRequest) => {
      req.deviceCodeCallback({
        message: 'To sign in ... enter code ABCD1234 ...',
        userCode: 'ABCD1234',
        verificationUri: 'https://microsoft.com/devicelogin',
        deviceCode: 'd',
        expiresIn: 900,
        interval: 5,
      });
      await pollingPromise;
      return {
        accessToken: 'tok',
        account: FAKE_ACCOUNT,
        expiresOn: new Date(Date.now() + 3_600_000),
      };
    });
    const app = { acquireTokenByDeviceCode: acquireMock } as unknown as PublicClientApplication;

    const t0 = Date.now();
    const result = await loginDeviceCodeInBackground({ app, cache });
    const elapsed = Date.now() - t0;

    // We must NOT have waited for MSAL polling to complete.
    expect(elapsed).toBeLessThan(200);
    expect(result.userCode).toBe('ABCD1234');
    expect(result.verificationUrl).toBe('https://microsoft.com/devicelogin');
    expect(typeof result.expiresAt).toBe('string');
    expect(new Date(result.expiresAt).getTime()).toBeGreaterThan(Date.now());

    // The background completion is still pending.
    release();
    const completion = await result.completion;
    expect(completion).toEqual({ ok: true, upn: 'a@x.com' });
  });

  it('completion resolves with {ok: false, reason} when MSAL rejects', async () => {
    const cache = new InMemoryTokenCache();
    const acquireMock = vi.fn(async (req: DeviceCodeRequest) => {
      req.deviceCodeCallback({
        message: 'msg',
        userCode: 'CODE',
        verificationUri: 'https://microsoft.com/devicelogin',
        deviceCode: 'd',
        expiresIn: 900,
        interval: 5,
      });
      throw new Error('user_timed_out');
    });
    const app = { acquireTokenByDeviceCode: acquireMock } as unknown as PublicClientApplication;

    const result = await loginDeviceCodeInBackground({ app, cache });
    const completion = await result.completion;
    expect(completion.ok).toBe(false);
    if (!completion.ok) expect(completion.reason).toContain('user_timed_out');
  });

  it('rejects synchronously if MSAL never emits a device code', async () => {
    const cache = new InMemoryTokenCache();
    const acquireMock = vi.fn(async () => {
      throw new Error('msal_unreachable');
    });
    const app = { acquireTokenByDeviceCode: acquireMock } as unknown as PublicClientApplication;

    await expect(loginDeviceCodeInBackground({ app, cache })).rejects.toThrow('msal_unreachable');
  });

  it('runs MSAL acquisition under cache.lock()', async () => {
    const cache = new InMemoryTokenCache();
    const lockSpy = vi.spyOn(cache, 'lock');
    const acquireMock = vi.fn(async (req: DeviceCodeRequest) => {
      req.deviceCodeCallback({
        message: 'msg',
        userCode: 'CODE',
        verificationUri: 'https://microsoft.com/devicelogin',
        deviceCode: 'd',
        expiresIn: 900,
        interval: 5,
      });
      return {
        accessToken: 'tok',
        account: FAKE_ACCOUNT,
        expiresOn: new Date(Date.now() + 3_600_000),
      };
    });
    const app = { acquireTokenByDeviceCode: acquireMock } as unknown as PublicClientApplication;
    const result = await loginDeviceCodeInBackground({ app, cache });
    await result.completion;
    expect(lockSpy).toHaveBeenCalledTimes(1);
  });
});
