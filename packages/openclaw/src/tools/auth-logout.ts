/**
 * `auth_logout` — wipe this agent's cached token.
 *
 * The next tool call will return `auth_cache_missing` until `auth_login` runs
 * again. Idempotent (clearing an already-empty cache is fine).
 */
import { Type } from 'typebox';
import { FileTokenCache, logout } from '@alavida-ai/outlook-core';

import { _resetClientForTesting, resolveCachePath } from '../client.js';
import { defineTool } from '../register.js';

interface AuthLogoutResult {
  status: 'logged_out';
  agentId: string | null;
  cachePath: string;
}

const authLogout = defineTool({
  name: 'outlook_auth_logout',
  description:
    "Clear this agent's cached Outlook token. The next tool call will require auth_login again.",
  parameters: Type.Object({}),
  async execute(_params, config): Promise<AuthLogoutResult> {
    const cachePath = resolveCachePath(config);
    const cache = new FileTokenCache(cachePath);
    await logout({ cache });
    // Drop the memoised OutlookClient so the next call rebuilds against an
    // empty cache (MSAL holds in-memory accounts otherwise).
    _resetClientForTesting();
    return {
      status: 'logged_out',
      agentId: config.agentId ?? null,
      cachePath,
    };
  },
});

export default authLogout;
