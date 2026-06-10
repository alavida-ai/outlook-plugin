/**
 * `auth_login` — start a device-code login for this agent.
 *
 * Fire-and-forget: returns the verification URL + 6-character code as soon as
 * Microsoft emits them (~200 ms). The plugin continues polling Microsoft in
 * the background; the agent surfaces URL+code to the human, waits for the
 * human to confirm, then calls `auth_status` to verify the cached token.
 *
 * Multi-agent isolation: tokens land at `<agentDir>/outlook-tokens.json` by
 * default (see `client.ts:resolveCachePath`).
 */
import { Type } from 'typebox';
import {
  buildMsalApp,
  FileTokenCache,
  loginDeviceCodeInBackground,
} from '@alavida-ai/outlook-core';

import { resolveCachePath } from '../client.js';
import { defineTool } from '../register.js';

interface AuthLoginResult {
  status: 'pending';
  verificationUrl: string;
  userCode: string;
  expiresAt: string;
  agentId: string | null;
  cachePath: string;
  hint: string;
}

const authLogin = defineTool({
  name: 'outlook_auth_login',
  description:
    'Start an OAuth device-code login for this agent. Returns the URL and ' +
    'six-character code immediately — surface them to the human and wait ' +
    'for them to confirm sign-in. Then call auth_status to verify.',
  parameters: Type.Object({}),
  async execute(_params, config): Promise<AuthLoginResult> {
    const cachePath = resolveCachePath(config);
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
      verificationUrl: result.verificationUrl,
      userCode: result.userCode,
      expiresAt: result.expiresAt,
      agentId: config.agentId ?? null,
      cachePath,
      hint:
        'Open the URL on any device, enter the code, sign in. Then call ' +
        'outlook.auth_status to confirm.',
    };
  },
});

export default authLogin;
