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
import { defineTool } from '../register.js';

interface BrowserAuthLoginResult {
  status: 'pending';
  flow: 'browser';
  authUrl: string;
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
  clientId?: string;
  tenantId?: string;
}

/**
 * Register the server-side pending flow for a minted auth URL and build the
 * agent-facing envelope. The PKCE verifier and nonce stay in the pending-flows
 * map — only the opaque `authUrl` is ever returned to the agent.
 *
 * Pure (modulo the injected `register`) so it can be tested without MSAL.
 */
export function startBrowserFlow(
  input: StartBrowserFlowInput,
  register: (flow: PendingFlow) => void = registerPendingFlow,
): BrowserAuthLoginResult {
  const { urlResult, redirectUri, cachePath, agentId, clientId, tenantId } = input;
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
  return {
    status: 'pending',
    flow: 'browser',
    authUrl: urlResult.authUrl,
    expiresAt: urlResult.expiresAt,
    agentId,
    cachePath,
    hint:
      'Open the URL in a browser and sign in. Then call outlook_auth_status ' +
      'to confirm.',
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
      clientId: config.clientId,
      tenantId: config.tenantId,
    });
  },
});

export default authLogin;
