import type { AccountInfo, PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import { resolveAccount } from './accounts.js';
import { AuthAmbiguousAccountError } from './errors.js';

export interface StatusInput {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn?: string;
}

/**
 * Return the cached account info if logged in, `null` if not.
 *
 * Multi-account ambiguity still throws — `outlook auth status` is a check,
 * not a workaround for picking among accounts. The caller passes
 * `preferredUpn` when they already know which mailbox they're asking about.
 */
export async function status(input: StatusInput): Promise<AccountInfo | null> {
  const { app, cache: _cache, preferredUpn } = input;
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length === 0) return null;
  try {
    return resolveAccount(accounts, preferredUpn);
  } catch (err) {
    if (err instanceof AuthAmbiguousAccountError) throw err;
    return null;
  }
}
