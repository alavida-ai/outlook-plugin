import type { AccountInfo } from '@azure/msal-node';

import { AuthAmbiguousAccountError, AuthCacheMissingError } from './errors.js';

/**
 * Pick a single cached account, applying the caller's preference if any.
 *
 * Per spec §4.2.5 — **never** silently pick `accounts[0]`. If the cache
 * holds more than one account and the caller hasn't disambiguated,
 * `AuthAmbiguousAccountError` lists the cached UPNs for the human/agent
 * to choose from.
 *
 * @param accounts MSAL's cached `AccountInfo[]` in cache order.
 * @param preferredUpn UPN to select. Undefined = "I have no preference".
 *                     Case-insensitive match.
 */
export function resolveAccount(
  accounts: readonly AccountInfo[],
  preferredUpn: string | undefined,
): AccountInfo {
  if (accounts.length === 0) {
    throw new AuthCacheMissingError();
  }

  if (preferredUpn !== undefined) {
    const needle = preferredUpn.toLowerCase();
    const hit = accounts.find((a) => a.username.toLowerCase() === needle);
    if (hit) return hit;
    throw new AuthAmbiguousAccountError(accounts.map((a) => a.username));
  }

  if (accounts.length === 1) {
    return accounts[0]!;
  }

  throw new AuthAmbiguousAccountError(accounts.map((a) => a.username));
}
