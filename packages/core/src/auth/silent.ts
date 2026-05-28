import type { AccountInfo, PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import { resolveAccount } from './accounts.js';
import {
  AuthInteractionRequiredError,
  AuthRefreshFailedError,
} from './errors.js';
import { OUTLOOK_SCOPES } from './msal.js';

export interface GetAccessTokenInput {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn: string | undefined;
}

export interface AccessTokenResult {
  accessToken: string;
  expiresOn: Date;
  account: AccountInfo;
}

/**
 * Acquire a valid Graph access token for the cached account.
 *
 * Holds the cache lock for the entire silent-refresh cycle so concurrent
 * CLI + plugin processes don't race against each other (spec §4.2.2). All
 * `AuthError` variants the caller might see are typed.
 */
export async function getAccessToken(input: GetAccessTokenInput): Promise<AccessTokenResult> {
  const { app, cache, preferredUpn } = input;
  return cache.lock(async () => {
    const accounts = await app.getTokenCache().getAllAccounts();
    const account = resolveAccount(accounts, preferredUpn);
    try {
      const result = await app.acquireTokenSilent({
        account,
        scopes: [...OUTLOOK_SCOPES],
      });
      if (!result?.accessToken || !result.expiresOn) {
        throw new AuthRefreshFailedError('MSAL returned no access token');
      }
      return {
        accessToken: result.accessToken,
        expiresOn: result.expiresOn,
        account,
      };
    } catch (err) {
      if (err instanceof AuthRefreshFailedError) throw err;
      if (isInteractionRequired(err)) {
        throw new AuthInteractionRequiredError((err as Error).message ?? 'interaction required');
      }
      throw new AuthRefreshFailedError(toMessage(err));
    }
  });
}

function isInteractionRequired(err: unknown): boolean {
  if (typeof err !== 'object' || err === null) return false;
  const code = (err as { errorCode?: string }).errorCode;
  return code === 'interaction_required' || code === 'consent_required' || code === 'login_required';
}

function toMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  return String(err);
}
