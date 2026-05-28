import { Client } from '@microsoft/microsoft-graph-client';
import type { PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from '../auth/cache.js';
import { MsalAuthenticationProvider } from './auth-provider.js';

export interface MakeGraphClientOptions {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn: string | undefined;
}

/**
 * Construct a configured `@microsoft/microsoft-graph-client` instance whose
 * auth provider delegates to our MSAL silent-refresh path.
 *
 * The SDK's middleware chain handles `Retry-After` on 429 / 503 out of the
 * box; we don't add custom middleware here.
 */
export function makeGraphClient(options: MakeGraphClientOptions): Client {
  const authProvider = new MsalAuthenticationProvider(options);
  return Client.initWithMiddleware({ authProvider });
}
