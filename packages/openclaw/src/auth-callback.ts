/**
 * Browser auth-callback route for the Authorization Code + PKCE flow.
 *
 * This module owns the server-side pending-flows map: the gate that ties a
 * Microsoft redirect back to the agent that started the login. `auth_login`
 * (browser variant) registers a flow keyed by a random `state`; Microsoft
 * redirects the user's browser to `/outlook/auth-callback?code=…&state=…`;
 * this handler validates `state`, redeems the `code` against the PKCE verifier,
 * and lets `exchangeAuthCode` write tokens to the initiating agent's own cache.
 *
 * Security invariants (see ALA-765 security model):
 *   - `state` is single-use: it is removed from the map the instant a callback
 *     claims it, so a replayed or double-submitted callback finds nothing.
 *   - expired flows (>10 min) are refused and dropped.
 *   - the PKCE `verifier`, the `code`, and tokens NEVER appear in any HTTP
 *     response body or log line.
 *   - per-agent isolation: the resolved cache path / client / tenant captured
 *     at flow start drive which cache the tokens land in.
 */
import type { IncomingMessage, ServerResponse } from 'node:http';
import {
  buildMsalApp,
  exchangeAuthCode,
  FileTokenCache,
} from '@alavida-ai/outlook-core';
import type { OpenClawPluginHttpRouteHandler } from 'openclaw/plugin-sdk/plugin-entry';

/** Public path this route is mounted at (also the Tailscale Funnel path). */
export const AUTH_CALLBACK_PATH = '/outlook/auth-callback';

/** Upper bound on the token-exchange leg before the handler gives up. */
export const AUTH_CALLBACK_TIMEOUT_MS = 30_000;

/**
 * One in-flight browser login. Everything the callback needs to finish the
 * flow for the right agent, captured when `auth_login` minted the URL.
 */
export interface PendingFlow {
  /** CSRF token; also the map key. */
  state: string;
  /** PKCE code verifier. Server-side only. */
  verifier: string;
  /** ID-token nonce we issued; validated on the ID token. */
  nonce: string;
  /** Redirect URI used to build the auth URL; must match on exchange. */
  redirectUri: string;
  /** Resolved per-agent token cache path the tokens must land in. */
  cachePath: string;
  /** Initiating agent id, for logging + isolation context. */
  agentId: string | null;
  clientId?: string;
  tenantId?: string;
  /** Epoch ms after which the flow is refused. */
  expiresAt: number;
}

/**
 * Redeem a claimed flow's code for tokens. Injectable so the route logic can be
 * tested without standing up MSAL or touching disk. Returns the signed-in UPN.
 */
export type ExchangeFn = (flow: PendingFlow, code: string) => Promise<{ upn: string }>;

interface AuthCallbackDeps {
  exchange?: ExchangeFn;
  now?: () => number;
}

// Module-level: shared between `auth_login` (writer) and the route (reader).
// A login is at most one entry per in-flight sign-in, so the map stays tiny.
const pendingFlows = new Map<string, PendingFlow>();

/** Register a freshly-minted browser login. Opportunistically GCs stale flows. */
export function registerPendingFlow(flow: PendingFlow): void {
  gcExpiredFlows();
  pendingFlows.set(flow.state, flow);
}

/** Test/diagnostic helper: number of pending flows currently held. */
export function pendingFlowCount(): number {
  return pendingFlows.size;
}

/** Test helper: drop all pending flows. */
export function clearPendingFlows(): void {
  pendingFlows.clear();
}

/** Remove flows whose TTL has passed. Safe to call often. */
export function gcExpiredFlows(now: number = Date.now()): void {
  for (const [state, flow] of pendingFlows) {
    if (now > flow.expiresAt) pendingFlows.delete(state);
  }
}

function defaultExchange(flow: PendingFlow, code: string): Promise<{ upn: string }> {
  const cache = new FileTokenCache(flow.cachePath);
  const app = buildMsalApp({
    cache,
    clientId: flow.clientId,
    tenantId: flow.tenantId,
  });
  return exchangeAuthCode({
    app,
    cache,
    code,
    verifier: flow.verifier,
    redirectUri: flow.redirectUri,
    expectedNonce: flow.nonce,
    expectedState: flow.state,
  }).then((result) => ({ upn: result.account.username }));
}

function htmlResponse(res: ServerResponse, status: number, title: string, body: string): true {
  const html =
    `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title>` +
    `<meta name="viewport" content="width=device-width, initial-scale=1">` +
    `<style>body{font-family:system-ui,sans-serif;max-width:32rem;margin:4rem auto;padding:0 1rem;text-align:center}` +
    `h1{font-size:1.4rem}p{color:#555}</style></head>` +
    `<body><h1>${title}</h1><p>${body}</p></body></html>`;
  res.writeHead(status, {
    'Content-Type': 'text/html; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  res.end(html);
  return true;
}

function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('token exchange timed out')), ms);
    promise.then(
      (v) => {
        clearTimeout(timer);
        resolve(v);
      },
      (err) => {
        clearTimeout(timer);
        reject(err);
      },
    );
  });
}

/**
 * Build the HTTP route handler. The default wires the real MSAL token exchange;
 * tests inject a fake `exchange`.
 */
export function makeAuthCallbackHandler(
  deps: AuthCallbackDeps = {},
): OpenClawPluginHttpRouteHandler {
  const exchange = deps.exchange ?? defaultExchange;
  const now = deps.now ?? (() => Date.now());

  return async function handleAuthCallback(
    req: IncomingMessage,
    res: ServerResponse,
  ): Promise<boolean> {
    const url = new URL(req.url ?? '', 'http://localhost');
    const state = url.searchParams.get('state') ?? '';
    const code = url.searchParams.get('code');
    const providerError = url.searchParams.get('error');

    // 1. Validate state against the server-side map. An unknown state is the
    //    CSRF / stale-link case — refuse without revealing anything.
    const flow = state ? pendingFlows.get(state) : undefined;
    if (!flow) {
      return htmlResponse(
        res,
        403,
        'Sign-in link expired',
        'This sign-in link is invalid or has already been used. Start a new sign-in and try again.',
      );
    }

    // 2. Single-use claim: remove now so a replay or double-submit finds nothing,
    //    even if the exchange below is still in flight.
    pendingFlows.delete(state);

    // 3. Expiry.
    if (now() > flow.expiresAt) {
      return htmlResponse(
        res,
        403,
        'Sign-in link expired',
        'This sign-in link has expired. Start a new sign-in and try again.',
      );
    }

    // 4. Provider-side failure (user cancelled, consent withheld, …).
    if (providerError) {
      console.error(
        `[outlook.auth-callback] provider error for agent=${flow.agentId ?? '<none>'}: ${providerError}`,
      );
      return htmlResponse(
        res,
        400,
        'Sign-in did not complete',
        'Microsoft reported that sign-in was cancelled or denied. You can close this tab and try again.',
      );
    }

    // 5. Must have a code to redeem.
    if (!code) {
      return htmlResponse(
        res,
        400,
        'Sign-in did not complete',
        'The sign-in response was missing its authorization code. Start a new sign-in and try again.',
      );
    }

    // 6. Redeem the code for tokens (per-agent cache write happens inside).
    try {
      await withTimeout(exchange(flow, code), AUTH_CALLBACK_TIMEOUT_MS);
      console.error(
        `[outlook.auth-callback] sign-in complete for agent=${flow.agentId ?? '<none>'}`,
      );
      return htmlResponse(
        res,
        200,
        'Sign-in complete',
        'You are signed in. You can close this tab and return to your conversation.',
      );
    } catch (err) {
      // Never echo the underlying error (it can carry request detail) to the
      // browser; log a redacted line server-side instead.
      const message = err instanceof Error ? err.message : String(err);
      console.error(
        `[outlook.auth-callback] token exchange failed for agent=${flow.agentId ?? '<none>'}: ${message}`,
      );
      return htmlResponse(
        res,
        500,
        'Sign-in could not be completed',
        'Something went wrong finishing sign-in. Start a new sign-in and try again.',
      );
    }
  };
}
