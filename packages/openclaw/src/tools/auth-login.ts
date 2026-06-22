/**
 * `auth_login` — start a browser (Authorization Code + PKCE) login for this agent.
 *
 * Requires `oauthRedirectUri` in plugin config (the public HTTPS callback URL).
 * Returns a sign-in URL immediately; the agent surfaces it to the human, who
 * signs in via the browser. Microsoft redirects to the plugin's
 * `/outlook/auth-callback` route, which redeems the code and writes tokens to
 * this agent's cache. The agent then calls `auth_status` to confirm.
 *
 * There is no device-code fallback: device-code is blocked by the Conditional
 * Access baselines we target, and the CLI's localhost interactive flow can't
 * run on a headless gateway — so the browser flow is the only gateway path.
 *
 * Multi-agent isolation: tokens land at `<agentDir>/outlook-tokens.json` by
 * default (see `client.ts:resolveCachePath`).
 */
import { Type } from 'typebox';
import {
  buildAuthCodeUrl,
  buildMsalApp,
  FileTokenCache,
  type AuthCodeUrlResult,
} from '@alavida-ai/outlook-core';

import { resolveCachePath } from '../client.js';
import { registerPendingFlow, type PendingFlow } from '../auth-callback.js';
import { AUTH_MESSAGE_TTL_MS, stashAuthMessage } from '../auth-message.js';
import { defineTool } from '../register.js';

interface BrowserAuthLoginResult {
  status: 'pending';
  flow: 'browser';
  /**
   * `channel` — the URL was sent to the user out-of-band (via the
   * message_sending hook) and is deliberately absent from this envelope so the
   * agent can't alter it. `inline` — no session context, so the URL is returned
   * here as a fallback (the hook can't deliver it).
   */
  delivery: 'channel' | 'inline';
  /** Present only when `delivery === 'inline'`. */
  authUrl?: string;
  expiresAt: string;
  agentId: string | null;
  cachePath: string;
  hint: string;
}

interface StartBrowserFlowInput {
  urlResult: AuthCodeUrlResult;
  redirectUri: string;
  cachePath: string;
  agentId: string | null;
  /** The conversation/session the link must be delivered to; null if unknown. */
  sessionKey: string | null;
  clientId?: string;
  tenantId?: string;
}

/**
 * Register the server-side pending flow for a minted auth URL and build the
 * agent-facing envelope.
 *
 * Security: when a `sessionKey` is present, the URL is **stashed** for the
 * message_sending hook to deliver to the user directly, and is kept entirely
 * out of the returned envelope — a prompt-injected agent never sees it and so
 * can't swap in a phishing link. Without a session (no channel to deliver to),
 * we fall back to returning the URL inline so sign-in still works.
 *
 * Pure (modulo the injected `register`/`stash`) so it can be tested without MSAL.
 */
export function startBrowserFlow(
  input: StartBrowserFlowInput,
  register: (flow: PendingFlow) => void = registerPendingFlow,
  stash: (sessionKey: string, url: string, expiresAt: number) => void = stashAuthMessage,
): BrowserAuthLoginResult {
  const { urlResult, redirectUri, cachePath, agentId, sessionKey, clientId, tenantId } = input;
  register({
    state: urlResult.state,
    verifier: urlResult.verifier,
    nonce: urlResult.nonce,
    redirectUri,
    cachePath,
    agentId,
    clientId,
    tenantId,
    expiresAt: new Date(urlResult.expiresAt).getTime(),
  });

  if (sessionKey) {
    stash(sessionKey, urlResult.authUrl, Date.now() + AUTH_MESSAGE_TTL_MS);
    return {
      status: 'pending',
      flow: 'browser',
      delivery: 'channel',
      expiresAt: urlResult.expiresAt,
      agentId,
      cachePath,
      hint:
        'The sign-in link has been sent to the user in this channel. Ask them ' +
        'to confirm once they have signed in, then call outlook_auth_status to verify.',
    };
  }

  return {
    status: 'pending',
    flow: 'browser',
    delivery: 'inline',
    authUrl: urlResult.authUrl,
    expiresAt: urlResult.expiresAt,
    agentId,
    cachePath,
    hint:
      'Surface this URL to the human and wait for sign-in. Then call ' +
      'outlook_auth_status to confirm.',
  };
}

const authLogin = defineTool({
  name: 'outlook_auth_login',
  description:
    'Start a browser sign-in for this agent. Returns a Microsoft sign-in URL — ' +
    'surface it to the human and wait for them to sign in, then call ' +
    'auth_status to verify. Requires oauthRedirectUri in plugin config.',
  parameters: Type.Object({}),
  async execute(_params, config): Promise<BrowserAuthLoginResult> {
    if (!config.oauthRedirectUri) {
      throw new Error(
        'outlook_auth_login requires `oauthRedirectUri` in the outlook plugin ' +
          'config — the public HTTPS callback URL for the browser sign-in flow ' +
          '(e.g. https://<gateway>.<tailnet>.ts.net/outlook/auth-callback). ' +
          'See the outlook skill auth reference for setup.',
      );
    }

    const cachePath = resolveCachePath(config);
    const cache = new FileTokenCache(cachePath);
    const app = buildMsalApp({
      cache,
      clientId: config.clientId,
      tenantId: config.tenantId,
    });
    const urlResult = await buildAuthCodeUrl({
      app,
      redirectUri: config.oauthRedirectUri,
    });
    return startBrowserFlow({
      urlResult,
      redirectUri: config.oauthRedirectUri,
      cachePath,
      agentId: config.agentId ?? null,
      sessionKey: config.sessionKey ?? null,
      clientId: config.clientId,
      tenantId: config.tenantId,
    });
  },
});

export default authLogin;
