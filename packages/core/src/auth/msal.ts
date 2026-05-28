import {
  PublicClientApplication,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from '@azure/msal-node';

import type { TokenCache } from './cache.js';

/**
 * The shared multi-tenant Entra app that ships with `outlook-cli`. Same id
 * the Python implementation used, so users migrating from Python don't
 * re-consent.
 */
export const EMBEDDED_CLIENT_ID = '18f9e6ff-2b0a-423e-bb35-ab9b541e604e';

/**
 * `common` accepts both personal Microsoft accounts and any work/school
 * tenant. MSAL resolves the actual tenant from the user's sign-in. Required
 * for multi-tenant apps.
 */
export const EMBEDDED_TENANT = 'common';

/**
 * Delegated scopes the CLI requests at sign-in. `offline_access` is added
 * implicitly by MSAL for public clients.
 */
export const OUTLOOK_SCOPES = [
  'Mail.ReadWrite',
  'Calendars.ReadWrite',
  'Calendars.ReadWrite.Shared',
  'Contacts.ReadWrite',
  'User.Read',
] as const;

export interface BuildMsalAppOptions {
  cache: TokenCache;
  clientId?: string;
  tenantId?: string;
}

/**
 * Construct an MSAL `PublicClientApplication` wired to our `TokenCache`.
 *
 * MSAL serialises its cache to a string blob; we treat that blob as opaque
 * and let `TokenCache` own atomicity, integrity, and locking (see §4.2).
 */
export function buildMsalApp(options: BuildMsalAppOptions): PublicClientApplication {
  const clientId = options.clientId ?? EMBEDDED_CLIENT_ID;
  const tenantId = options.tenantId ?? EMBEDDED_TENANT;

  const cachePlugin: ICachePlugin = {
    async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
      const blob = await options.cache.load().catch((err) => {
        // Corrupt cache: surface as "empty" to MSAL; the AuthCacheCorruptError
        // is re-raised by the upper-layer getAccessToken on the next user op.
        ctx.tokenCache.deserialize('');
        throw err;
      });
      if (blob) ctx.tokenCache.deserialize(blob);
    },
    async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
      if (!ctx.cacheHasChanged) return;
      await options.cache.save(ctx.tokenCache.serialize());
    },
  };

  const config: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: { cachePlugin },
  };
  return new PublicClientApplication(config);
}
