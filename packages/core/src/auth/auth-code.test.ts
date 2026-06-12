import { describe, expect, it, vi } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-node';

import { InMemoryTokenCache } from './cache.js';
import { AuthRefreshFailedError } from './errors.js';
import {
  buildAuthCodeUrl,
  exchangeAuthCode,
  PENDING_FLOW_TTL_MS,
} from './auth-code.js';

const FAKE_ACCOUNT = {
  homeAccountId: 'h',
  environment: 'e',
  tenantId: 't',
  username: 'a@x.com',
  localAccountId: 'l',
};

describe('buildAuthCodeUrl', () => {
  it('returns the URL minted by MSAL plus a state, verifier, nonce and expiry', async () => {
    const getAuthCodeUrl = vi.fn(async () => 'https://login.microsoftonline.com/authorize?x=1');
    const app = { getAuthCodeUrl } as unknown as PublicClientApplication;

    const result = await buildAuthCodeUrl({
      app,
      redirectUri: 'https://gw.ts.net/outlook/auth-callback',
    });

    expect(result.authUrl).toBe('https://login.microsoftonline.com/authorize?x=1');
    expect(result.state).toMatch(/.+/);
    expect(result.verifier).toMatch(/.+/);
    expect(result.nonce).toMatch(/.+/);
    expect(typeof result.expiresAt).toBe('string');
    expect(new Date(result.expiresAt).getTime()).toBeGreaterThan(Date.now());
  });

  it('passes PKCE S256 challenge, state, nonce, redirectUri and scopes to MSAL', async () => {
    const getAuthCodeUrl = vi.fn(async () => 'https://example/authorize');
    const app = { getAuthCodeUrl } as unknown as PublicClientApplication;

    const result = await buildAuthCodeUrl({
      app,
      redirectUri: 'https://gw.ts.net/outlook/auth-callback',
    });

    expect(getAuthCodeUrl).toHaveBeenCalledTimes(1);
    const req = getAuthCodeUrl.mock.calls[0][0] as Record<string, unknown>;
    expect(req.redirectUri).toBe('https://gw.ts.net/outlook/auth-callback');
    expect(req.codeChallengeMethod).toBe('S256');
    expect(typeof req.codeChallenge).toBe('string');
    expect((req.codeChallenge as string).length).toBeGreaterThan(0);
    // The challenge sent to Microsoft must not be the raw verifier.
    expect(req.codeChallenge).not.toBe(result.verifier);
    expect(req.state).toBe(result.state);
    expect(req.nonce).toBe(result.nonce);
    expect(req.scopes).toEqual(['Mail.ReadWrite', 'Calendars.Read', 'Calendars.Read.Shared', 'User.Read']);
  });

  it('sets expiry roughly PENDING_FLOW_TTL_MS into the future', async () => {
    const app = {
      getAuthCodeUrl: vi.fn(async () => 'https://example/authorize'),
    } as unknown as PublicClientApplication;

    const before = Date.now();
    const result = await buildAuthCodeUrl({
      app,
      redirectUri: 'https://gw.ts.net/outlook/auth-callback',
    });
    const ttl = new Date(result.expiresAt).getTime() - before;

    expect(ttl).toBeGreaterThan(PENDING_FLOW_TTL_MS - 2_000);
    expect(ttl).toBeLessThanOrEqual(PENDING_FLOW_TTL_MS + 1_000);
  });

  it('mints a distinct state and nonce on each call', async () => {
    const app = {
      getAuthCodeUrl: vi.fn(async () => 'https://example/authorize'),
    } as unknown as PublicClientApplication;

    const a = await buildAuthCodeUrl({ app, redirectUri: 'https://gw/cb' });
    const b = await buildAuthCodeUrl({ app, redirectUri: 'https://gw/cb' });

    expect(a.state).not.toBe(b.state);
    expect(a.nonce).not.toBe(b.nonce);
    expect(a.verifier).not.toBe(b.verifier);
    expect(a.state).not.toBe(a.nonce);
  });
});

describe('exchangeAuthCode', () => {
  function makeApp(overrides?: Partial<Record<string, unknown>>) {
    const acquireTokenByCode = vi.fn(async () => ({
      accessToken: 'tok',
      account: FAKE_ACCOUNT,
      expiresOn: new Date(Date.now() + 3_600_000),
      idTokenClaims: { nonce: 'NONCE-1' },
      ...overrides,
    }));
    return {
      app: { acquireTokenByCode } as unknown as PublicClientApplication,
      acquireTokenByCode,
    };
  }

  it('exchanges the code and returns the account + expiry', async () => {
    const cache = new InMemoryTokenCache();
    const { app, acquireTokenByCode } = makeApp();

    const result = await exchangeAuthCode({
      app,
      cache,
      code: 'AUTH_CODE',
      verifier: 'VERIFIER',
      redirectUri: 'https://gw/cb',
      expectedNonce: 'NONCE-1',
      expectedState: 'STATE-1',
    });

    expect(result.account.username).toBe('a@x.com');
    expect(result.expiresOn).toBeInstanceOf(Date);
    const req = acquireTokenByCode.mock.calls[0][0] as Record<string, unknown>;
    expect(req.code).toBe('AUTH_CODE');
    expect(req.codeVerifier).toBe('VERIFIER');
    expect(req.redirectUri).toBe('https://gw/cb');
    expect(req.state).toBe('STATE-1');
  });

  it('runs the exchange under cache.lock()', async () => {
    const cache = new InMemoryTokenCache();
    const lockSpy = vi.spyOn(cache, 'lock');
    const { app } = makeApp();

    await exchangeAuthCode({
      app,
      cache,
      code: 'c',
      verifier: 'v',
      redirectUri: 'https://gw/cb',
      expectedNonce: 'NONCE-1',
    });

    expect(lockSpy).toHaveBeenCalledTimes(1);
  });

  it('throws AuthRefreshFailedError when MSAL returns no account', async () => {
    const cache = new InMemoryTokenCache();
    const { app } = makeApp({ account: null });

    await expect(
      exchangeAuthCode({
        app,
        cache,
        code: 'c',
        verifier: 'v',
        redirectUri: 'https://gw/cb',
        expectedNonce: 'NONCE-1',
      }),
    ).rejects.toBeInstanceOf(AuthRefreshFailedError);
  });

  it('rejects an ID token whose nonce does not match the one we issued', async () => {
    const cache = new InMemoryTokenCache();
    const { app } = makeApp({ idTokenClaims: { nonce: 'ATTACKER-NONCE' } });

    await expect(
      exchangeAuthCode({
        app,
        cache,
        code: 'c',
        verifier: 'v',
        redirectUri: 'https://gw/cb',
        expectedNonce: 'NONCE-1',
      }),
    ).rejects.toBeInstanceOf(AuthRefreshFailedError);
  });
});
