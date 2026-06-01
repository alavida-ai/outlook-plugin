/**
 * `auth_status` — return the signed-in account for this agent.
 *
 * Companion to `auth_login`. The agent calls this after the human says
 * sign-in completed (or speculatively, to discover whether this agent is
 * already authenticated).
 *
 * Strategy: first check the cached MSAL account (no Graph round-trip). If
 * absent, return `not_authenticated` immediately. If present, do a silent
 * refresh to validate the token's still good. Auth-layer errors from the
 * refresh map to the same `not_authenticated` envelope so the agent always
 * gets a clean two-state answer.
 *
 * Other errors (network, throttling, server) bubble up via
 * `withErrorMapping` and surface as `__toolError` envelopes.
 */
import { Type } from 'typebox';
import {
  AuthAmbiguousAccountError,
  AuthCacheCorruptError,
  AuthCacheMissingError,
  AuthInteractionRequiredError,
  AuthRefreshFailedError,
  FileTokenCache,
  buildMsalApp,
  getAccessToken,
  status as cachedStatus,
} from '@alavida-ai/outlook-core';

import { resolveCachePath } from '../client.js';
import { defineTool } from '../register.js';

type AuthStatusResult =
  | {
      status: 'authenticated';
      upn: string;
      displayName: string | null;
      agentId: string | null;
      cachePath: string;
    }
  | {
      status: 'not_authenticated';
      reason:
        | 'no_cache'
        | 'cache_corrupt'
        | 'refresh_failed'
        | 'interactive_required'
        | 'ambiguous_account';
      agentId: string | null;
      cachePath: string;
      hint: string;
      accounts?: readonly string[];
    };

const authStatus = defineTool({
  name: 'auth_status',
  description:
    'Report whether this agent has a usable cached token. After auth_login ' +
    'and the human confirms sign-in, call this to confirm. Returns ' +
    'authenticated | not_authenticated; auth errors map to not_authenticated.',
  parameters: Type.Object({}),
  async execute(_params, config): Promise<AuthStatusResult> {
    const cachePath = resolveCachePath(config);
    const cache = new FileTokenCache(cachePath);
    const app = buildMsalApp({
      cache,
      clientId: config.clientId,
      tenantId: config.tenantId,
    });

    // Step 1: cheap cache-only check — never touches Graph.
    let account;
    try {
      account = await cachedStatus({ app, cache, preferredUpn: config.account });
    } catch (err) {
      if (err instanceof AuthAmbiguousAccountError) {
        return {
          status: 'not_authenticated',
          reason: 'ambiguous_account',
          agentId: config.agentId ?? null,
          cachePath,
          hint:
            'Multiple accounts cached. Set `account: "<upn>"` in plugin config to pin one.',
          accounts: err.accounts,
        };
      }
      if (err instanceof AuthCacheCorruptError) {
        return {
          status: 'not_authenticated',
          reason: 'cache_corrupt',
          agentId: config.agentId ?? null,
          cachePath,
          hint: 'Token cache is unreadable. Call outlook.auth_logout then outlook.auth_login.',
        };
      }
      throw err;
    }

    if (account === null) {
      return {
        status: 'not_authenticated',
        reason: 'no_cache',
        agentId: config.agentId ?? null,
        cachePath,
        hint: 'Call outlook.auth_login.',
      };
    }

    // Step 2: silent refresh — checks the token isn't revoked.
    try {
      await getAccessToken({ app, cache, preferredUpn: config.account });
    } catch (err) {
      if (err instanceof AuthCacheMissingError) {
        return {
          status: 'not_authenticated',
          reason: 'no_cache',
          agentId: config.agentId ?? null,
          cachePath,
          hint: 'Call outlook.auth_login.',
        };
      }
      if (err instanceof AuthRefreshFailedError) {
        return {
          status: 'not_authenticated',
          reason: 'refresh_failed',
          agentId: config.agentId ?? null,
          cachePath,
          hint:
            'Cached token is stale (password change, revoked consent, …). Call outlook.auth_login.',
        };
      }
      if (err instanceof AuthInteractionRequiredError) {
        return {
          status: 'not_authenticated',
          reason: 'interactive_required',
          agentId: config.agentId ?? null,
          cachePath,
          hint: 'Microsoft requires interactive sign-in. Call outlook.auth_login.',
        };
      }
      throw err;
    }

    return {
      status: 'authenticated',
      upn: account.username,
      displayName: account.name ?? null,
      agentId: config.agentId ?? null,
      cachePath,
    };
  },
});

export default authStatus;
