import type { PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import type { LoginResult } from './login-result.js';
import { AuthRefreshFailedError } from './errors.js';
import { OUTLOOK_SCOPES } from './msal.js';

export interface LoginInteractiveInput {
  app: PublicClientApplication;
  cache: TokenCache;
  /**
   * Launch the system browser at `url`. MSAL stands up a localhost loopback
   * server to catch the redirect; Microsoft accepts `http://localhost` without
   * TLS for public clients, so no redirect URI needs registering.
   */
  openBrowser: (url: string) => Promise<void>;
  /** HTML shown in the browser tab on success. */
  successTemplate?: string;
  /** HTML shown in the browser tab on failure. */
  errorTemplate?: string;
}

/**
 * Run the interactive Authorization-Code flow for a human at a terminal:
 * open the browser, let MSAL's loopback server catch the redirect, and write
 * the resulting tokens through the cache plugin.
 *
 * This is the CLI counterpart to the OpenClaw plugin's browser flow — same
 * modern flow Microsoft recommends for public clients, but using a localhost
 * loopback instead of a public Funnel URL, so it works with no infra and is
 * compatible with Conditional Access policies that block device-code.
 *
 * Holds the cache lock for the full acquisition (spec §4.2.2).
 */
export async function loginInteractive(input: LoginInteractiveInput): Promise<LoginResult> {
  const { app, cache, openBrowser, successTemplate, errorTemplate } = input;
  return cache.lock(async () => {
    const result = await app.acquireTokenInteractive({
      scopes: [...OUTLOOK_SCOPES],
      openBrowser,
      successTemplate,
      errorTemplate,
    });
    if (!result?.account || !result.expiresOn) {
      throw new AuthRefreshFailedError('interactive flow returned no account');
    }
    return { account: result.account, expiresOn: result.expiresOn };
  });
}
