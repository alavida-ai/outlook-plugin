import type { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import type { PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from '../auth/cache.js';
import { getAccessToken } from '../auth/silent.js';

export interface MsalAuthenticationProviderOptions {
  app: PublicClientApplication;
  cache: TokenCache;
  preferredUpn: string | undefined;
}

/**
 * `AuthenticationProvider` for `@microsoft/microsoft-graph-client` that
 * delegates to our MSAL silent-refresh path. The Graph SDK calls
 * `getAccessToken()` once per request (modulo its own caching).
 */
export class MsalAuthenticationProvider implements AuthenticationProvider {
  constructor(private readonly options: MsalAuthenticationProviderOptions) {}

  async getAccessToken(): Promise<string> {
    const result = await getAccessToken(this.options);
    return result.accessToken;
  }
}
