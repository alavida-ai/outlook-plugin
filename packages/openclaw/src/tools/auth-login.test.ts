/**
 * Browser-flow wiring for auth_login. The MSAL URL assembly itself is covered
 * by core's buildAuthCodeUrl tests; here we verify the tool maps a minted URL
 * result + per-agent config into (a) the right pending-flow entry and (b) the
 * agent-facing envelope, without touching MSAL or the network.
 */
import type { AuthCodeUrlResult } from '@alavida-ai/outlook-core';
import { describe, expect, it, vi } from 'vitest';

import type { PendingFlow } from '../auth-callback.js';
import { startBrowserFlow } from './auth-login.js';

const URL_RESULT: AuthCodeUrlResult = {
  authUrl: 'https://login.microsoftonline.com/authorize?state=STATE-1',
  state: 'STATE-1',
  verifier: 'VERIFIER-1',
  nonce: 'NONCE-1',
  expiresAt: new Date(Date.now() + 600_000).toISOString(),
};

describe('startBrowserFlow', () => {
  it('returns a pending browser envelope carrying the auth URL', () => {
    const register = vi.fn<(f: PendingFlow) => void>();
    const result = startBrowserFlow(
      {
        urlResult: URL_RESULT,
        redirectUri: 'https://gw.ts.net/outlook/auth-callback',
        cachePath: '/agents/alfred/outlook-tokens.json',
        agentId: 'alfred',
      },
      register,
    );

    expect(result.status).toBe('pending');
    expect(result.flow).toBe('browser');
    expect(result.authUrl).toBe(URL_RESULT.authUrl);
    expect(result.expiresAt).toBe(URL_RESULT.expiresAt);
    expect(result.agentId).toBe('alfred');
    expect(result.cachePath).toBe('/agents/alfred/outlook-tokens.json');
    expect(result.hint).toMatch(/auth_status/);
  });

  it('registers a single-use pending flow holding the server-side secrets', () => {
    const captured: PendingFlow[] = [];
    const register = (f: PendingFlow) => captured.push(f);

    startBrowserFlow(
      {
        urlResult: URL_RESULT,
        redirectUri: 'https://gw.ts.net/outlook/auth-callback',
        cachePath: '/agents/alfred/outlook-tokens.json',
        agentId: 'alfred',
        clientId: 'cid',
        tenantId: 'tid',
      },
      register,
    );

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
    // Stored as epoch ms matching the issued ISO expiry.
    expect(captured[0].expiresAt).toBe(new Date(URL_RESULT.expiresAt).getTime());
  });

  it('never exposes the PKCE verifier or nonce in the agent-facing envelope', () => {
    const register = vi.fn<(f: PendingFlow) => void>();
    const result = startBrowserFlow(
      {
        urlResult: URL_RESULT,
        redirectUri: 'https://gw/cb',
        cachePath: '/c',
        agentId: null,
      },
      register,
    );

    const serialized = JSON.stringify(result);
    expect(serialized).not.toContain('VERIFIER-1');
    expect(serialized).not.toContain('NONCE-1');
  });
});
