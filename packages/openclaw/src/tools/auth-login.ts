/**
 * `auth_login` — start a login for this agent. Two flows, chosen by config:
 *
 *   - **Browser (Authorization Code + PKCE)** — when `oauthRedirectUri` is set
 *     in plugin config. Returns a sign-in URL immediately; the agent surfaces
 *     it to the human, who signs in via the browser. Microsoft redirects to the
 *     plugin's `/outlook/auth-callback` route, which redeems the code and writes
 *     tokens to this agent's cache. Required for tenants whose Conditional
 *     Access blocks device-code flow.
 *   - **Device code** — fallback when `oauthRedirectUri` is unset. Returns the
 *     verification URL + 6-character code; the plugin polls Microsoft in the
 *     background.
 *
 * Both are fire-and-forget: the agent surfaces what's returned, waits for the
 * human to confirm, then calls `auth_status` to verify the cached token.
 *
 * Multi-agent isolation: tokens land at `<agentDir>/outlook-tokens.json` by
 * default (see `client.ts:resolveCachePath`).
 */
import { Type } from 'typebox';
import {
  buildAuthCodeUrl,
  buildMsalApp,
  FileTokenCache,
  loginDeviceCodeInBackground,
  type AuthCodeUrlResult,
} from '@alavida-ai/outlook-core';

import { resolveCachePath, type PluginConfig } from '../client.js';
import { registerPendingFlow, type PendingFlow } from '../auth-callback.js';
import { defineTool } from '../register.js';

interface DeviceAuthLoginResult {
  status: 'pending';
  flow: 'device';
  verificationUrl: string;
  userCode: string;
  expiresAt: string;
  agentId: string | null;
  cachePath: string;
  hint: string;
}

interface BrowserAuthLoginResult {
  status: 'pending';
  flow: 'browser';
  authUrl: string;
  expiresAt: string;
  agentId: string | null;
  cachePath: string;
  hint: string;
}

type AuthLoginResult = DeviceAuthLoginResult | BrowserAuthLoginResult;

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

async function startDeviceFlow(
  config: PluginConfig,
  cachePath: string,
): Promise<DeviceAuthLoginResult> {
  const cache = new FileTokenCache(cachePath);
  const app = buildMsalApp({
    cache,
    clientId: config.clientId,
    tenantId: config.tenantId,
  });

  const result = await loginDeviceCodeInBackground({ app, cache });

  // Fire-and-forget: drain the completion promise so failures don't surface
  // as unhandled rejections. We never let them throw upstream — the agent
  // discovers outcome via the next auth_status call.
  void result.completion.then((outcome) => {
    if (!outcome.ok) {
      console.error(
        `[outlook.auth_login] device-code completion failed for agent=${config.agentId ?? '<none>'}: ${outcome.reason}`,
      );
    }
  });

  return {
    status: 'pending',
    flow: 'device',
    verificationUrl: result.verificationUrl,
    userCode: result.userCode,
    expiresAt: result.expiresAt,
    agentId: config.agentId ?? null,
    cachePath,
    hint:
      'Open the URL on any device, enter the code, sign in. Then call ' +
      'outlook_auth_status to confirm.',
  };
}

const authLogin = defineTool({
  name: 'outlook_auth_login',
  description:
    'Start an OAuth login for this agent. With a browser redirect configured, ' +
    'returns a sign-in URL; otherwise returns a device-code URL + six-character ' +
    'code. Surface what is returned to the human and wait for them to confirm ' +
    'sign-in, then call auth_status to verify.',
  parameters: Type.Object({}),
  async execute(_params, config): Promise<AuthLoginResult> {
    const cachePath = resolveCachePath(config);

    if (config.oauthRedirectUri) {
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
    }

    return startDeviceFlow(config, cachePath);
  },
});

export default authLogin;
