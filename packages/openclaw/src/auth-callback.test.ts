/**
 * The browser auth-callback route is the gate that ties a Microsoft redirect
 * back to the right agent's token cache. These tests exercise the security
 * invariants in isolation by injecting a fake token-exchange dependency:
 *   - unknown / expired / replayed `state` is refused without exchanging
 *   - a valid first hit exchanges exactly once and is single-use
 *   - provider errors and exchange failures never leak secrets
 */
import type { IncomingMessage, ServerResponse } from 'node:http';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
  makeAuthCallbackHandler,
  registerPendingFlow,
  clearPendingFlows,
  pendingFlowCount,
  gcExpiredFlows,
  type PendingFlow,
} from './auth-callback.js';

function flow(overrides: Partial<PendingFlow> = {}): PendingFlow {
  return {
    state: 'STATE-1',
    verifier: 'VERIFIER-1',
    nonce: 'NONCE-1',
    redirectUri: 'https://gw.ts.net/outlook/auth-callback',
    cachePath: '/tmp/agent/outlook-tokens.json',
    agentId: 'alfred',
    clientId: undefined,
    tenantId: undefined,
    expiresAt: Date.now() + 60_000,
    ...overrides,
  };
}

interface FakeRes {
  res: ServerResponse;
  statusCode: () => number;
  body: () => string;
  headers: () => Record<string, string | number>;
}

function fakeRes(): FakeRes {
  let status = 0;
  let chunk = '';
  const hdrs: Record<string, string | number> = {};
  const res = {
    writeHead(code: number, headers?: Record<string, string | number>) {
      status = code;
      Object.assign(hdrs, headers ?? {});
      return res;
    },
    setHeader(k: string, v: string | number) {
      hdrs[k] = v;
    },
    end(text?: string) {
      if (text) chunk += text;
      return res;
    },
  } as unknown as ServerResponse;
  return { res, statusCode: () => status, body: () => chunk, headers: () => hdrs };
}

function req(url: string): IncomingMessage {
  return { url, method: 'GET' } as IncomingMessage;
}

afterEach(() => {
  clearPendingFlows();
});

describe('auth-callback route handler', () => {
  it('exchanges a valid state+code exactly once and reports success', async () => {
    const exchange = vi.fn(async () => ({ upn: 'amit@sunglobal.com' }));
    const handler = makeAuthCallbackHandler({ exchange });
    registerPendingFlow(flow());

    const r = fakeRes();
    const handled = await handler(req('/outlook/auth-callback?state=STATE-1&code=AUTH_CODE'), r.res);

    expect(handled).toBe(true);
    expect(exchange).toHaveBeenCalledTimes(1);
    const [passedFlow, passedCode] = exchange.mock.calls[0];
    expect(passedFlow.verifier).toBe('VERIFIER-1');
    expect(passedFlow.nonce).toBe('NONCE-1');
    expect(passedCode).toBe('AUTH_CODE');
    expect(r.statusCode()).toBe(200);
    expect(String(r.headers()['Content-Type'])).toMatch(/text\/html/);
    // Flow consumed.
    expect(pendingFlowCount()).toBe(0);
  });

  it('refuses an unknown state without calling exchange', async () => {
    const exchange = vi.fn(async () => ({ upn: 'x' }));
    const handler = makeAuthCallbackHandler({ exchange });
    registerPendingFlow(flow());

    const r = fakeRes();
    await handler(req('/outlook/auth-callback?state=WRONG&code=AUTH_CODE'), r.res);

    expect(exchange).not.toHaveBeenCalled();
    expect(r.statusCode()).toBe(403);
    // The real pending flow is untouched.
    expect(pendingFlowCount()).toBe(1);
  });

  it('refuses an expired state and drops it without exchanging', async () => {
    const exchange = vi.fn(async () => ({ upn: 'x' }));
    const handler = makeAuthCallbackHandler({ exchange });
    registerPendingFlow(flow({ expiresAt: Date.now() - 1 }));

    const r = fakeRes();
    await handler(req('/outlook/auth-callback?state=STATE-1&code=AUTH_CODE'), r.res);

    expect(exchange).not.toHaveBeenCalled();
    expect(r.statusCode()).toBe(403);
    expect(pendingFlowCount()).toBe(0);
  });

  it('is single-use: a replayed state is refused on the second hit', async () => {
    const exchange = vi.fn(async () => ({ upn: 'amit@sunglobal.com' }));
    const handler = makeAuthCallbackHandler({ exchange });
    registerPendingFlow(flow());

    const r1 = fakeRes();
    await handler(req('/outlook/auth-callback?state=STATE-1&code=CODE-A'), r1.res);
    const r2 = fakeRes();
    await handler(req('/outlook/auth-callback?state=STATE-1&code=CODE-A'), r2.res);

    expect(exchange).toHaveBeenCalledTimes(1);
    expect(r1.statusCode()).toBe(200);
    expect(r2.statusCode()).toBe(403);
  });

  it('handles a provider error redirect without exchanging', async () => {
    const exchange = vi.fn(async () => ({ upn: 'x' }));
    const handler = makeAuthCallbackHandler({ exchange });
    registerPendingFlow(flow());

    const r = fakeRes();
    await handler(
      req('/outlook/auth-callback?state=STATE-1&error=access_denied&error_description=nope'),
      r.res,
    );

    expect(exchange).not.toHaveBeenCalled();
    expect(r.statusCode()).toBeGreaterThanOrEqual(400);
    // A failed flow is cleared so it can't be retried with a stale verifier.
    expect(pendingFlowCount()).toBe(0);
  });

  it('rejects a callback with no code and no error', async () => {
    const exchange = vi.fn(async () => ({ upn: 'x' }));
    const handler = makeAuthCallbackHandler({ exchange });
    registerPendingFlow(flow());

    const r = fakeRes();
    await handler(req('/outlook/auth-callback?state=STATE-1'), r.res);

    expect(exchange).not.toHaveBeenCalled();
    expect(r.statusCode()).toBeGreaterThanOrEqual(400);
  });

  it('surfaces an exchange failure as an error page without leaking the verifier or code', async () => {
    const exchange = vi.fn(async () => {
      throw new Error('AADSTS70008: expired');
    });
    const handler = makeAuthCallbackHandler({ exchange });
    registerPendingFlow(flow());

    const r = fakeRes();
    const handled = await handler(req('/outlook/auth-callback?state=STATE-1&code=AUTH_CODE'), r.res);

    expect(handled).toBe(true);
    expect(r.statusCode()).toBeGreaterThanOrEqual(400);
    expect(r.body()).not.toContain('VERIFIER-1');
    expect(r.body()).not.toContain('AUTH_CODE');
    // A failed exchange consumes the flow (single-use claim happens up front).
    expect(pendingFlowCount()).toBe(0);
  });
});

describe('pending-flows housekeeping', () => {
  it('gcExpiredFlows drops only expired entries', () => {
    registerPendingFlow(flow({ state: 'live', expiresAt: Date.now() + 60_000 }));
    registerPendingFlow(flow({ state: 'dead', expiresAt: Date.now() - 1 }));

    gcExpiredFlows();

    expect(pendingFlowCount()).toBe(1);
  });
});
