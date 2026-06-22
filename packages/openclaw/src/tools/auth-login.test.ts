/**
 * Browser-flow wiring for auth_login. The MSAL URL assembly itself is covered
 * by core's buildAuthCodeUrl tests; here we verify the tool maps a minted URL
 * result + per-agent config into (a) the right pending-flow entry, (b) the
 * out-of-band URL stash keyed by session, and (c) the agent-facing envelope —
 * which must NOT carry the URL when a session is present (so the agent can't
 * swap it for a phishing link).
 */
import type { AuthCodeUrlResult } from '@alavida-ai/outlook-core';
import { describe, expect, it, vi } from 'vitest';

import type { PendingFlow } from '../auth-callback.js';
import authLogin, { startBrowserFlow } from './auth-login.js';

describe('outlook_auth_login — browser-only (requires oauthRedirectUri)', () => {
  it('throws a clear error when oauthRedirectUri is not configured', async () => {
    await expect(
      authLogin.execute({}, { agentId: 'alfred', agentDir: '/tmp/agent' }),
    ).rejects.toThrow(/oauthRedirectUri/);
  });
});

const URL_RESULT: AuthCodeUrlResult = {
  authUrl: 'https://login.microsoftonline.com/authorize?state=STATE-1',
  state: 'STATE-1',
  verifier: 'VERIFIER-1',
  nonce: 'NONCE-1',
  expiresAt: new Date(Date.now() + 600_000).toISOString(),
};

function baseInput(extra: Record<string, unknown> = {}) {
  return {
    urlResult: URL_RESULT,
    redirectUri: 'https://gw.ts.net/outlook/auth-callback',
    cachePath: '/agents/alfred/outlook-tokens.json',
    agentId: 'alfred',
    sessionKey: 'sess-1',
    ...extra,
  };
}

describe('startBrowserFlow — with a session (out-of-band delivery)', () => {
  it('stashes the URL keyed by sessionKey and keeps it OUT of the agent envelope', () => {
    const register = vi.fn<(f: PendingFlow) => void>();
    const stash = vi.fn<(sessionKey: string, url: string, expiresAt: number) => void>();

    const result = startBrowserFlow(baseInput(), register, stash);

    // Stashed for the hook to deliver.
    expect(stash).toHaveBeenCalledTimes(1);
    const [sessionKey, url, expiresAt] = stash.mock.calls[0];
    expect(sessionKey).toBe('sess-1');
    expect(url).toBe(URL_RESULT.authUrl);
    expect(expiresAt).toBeGreaterThan(Date.now());

    // Agent-facing envelope: pending, channel delivery, NO url.
    expect(result.status).toBe('pending');
    expect(result.delivery).toBe('channel');
    expect(result.authUrl).toBeUndefined();
    expect(result.hint).toMatch(/auth_status/);
    // The URL never appears anywhere the agent can see.
    expect(JSON.stringify(result)).not.toContain('login.microsoftonline.com');
  });

  it('still registers the single-use pending flow with the server-side secrets', () => {
    const captured: PendingFlow[] = [];
    const register = (f: PendingFlow) => captured.push(f);

    startBrowserFlow(baseInput({ clientId: 'cid', tenantId: 'tid' }), register, vi.fn());

    expect(captured).toHaveLength(1);
    expect(captured[0]).toMatchObject({
      state: 'STATE-1',
      verifier: 'VERIFIER-1',
      nonce: 'NONCE-1',
      redirectUri: 'https://gw.ts.net/outlook/auth-callback',
      cachePath: '/agents/alfred/outlook-tokens.json',
      agentId: 'alfred',
      clientId: 'cid',
      tenantId: 'tid',
    });
    expect(captured[0].expiresAt).toBe(new Date(URL_RESULT.expiresAt).getTime());
  });

  it('never exposes the PKCE verifier, nonce, or URL in the agent envelope', () => {
    const result = startBrowserFlow(baseInput(), vi.fn(), vi.fn());
    const serialized = JSON.stringify(result);
    expect(serialized).not.toContain('VERIFIER-1');
    expect(serialized).not.toContain('NONCE-1');
    expect(serialized).not.toContain('login.microsoftonline.com');
  });
});

describe('startBrowserFlow — no session (inline fallback)', () => {
  it('returns the URL inline when there is no sessionKey (hook cannot deliver)', () => {
    const stash = vi.fn();
    const result = startBrowserFlow(baseInput({ sessionKey: null }), vi.fn(), stash);

    expect(stash).not.toHaveBeenCalled();
    expect(result.delivery).toBe('inline');
    expect(result.authUrl).toBe(URL_RESULT.authUrl);
    expect(result.hint).toMatch(/auth_status/);
  });
});
