import { describe, expect, it } from 'vitest';

import { InMemoryTokenCache } from './cache.js';
import { buildMsalApp, EMBEDDED_CLIENT_ID, EMBEDDED_TENANT, OUTLOOK_SCOPES } from './msal.js';

describe('buildMsalApp', () => {
  it('returns a PublicClientApplication wired to the embedded app id by default', () => {
    const app = buildMsalApp({ cache: new InMemoryTokenCache() });
    expect(app).toBeDefined();
    expect(EMBEDDED_CLIENT_ID).toBe('18f9e6ff-2b0a-423e-bb35-ab9b541e604e');
    expect(EMBEDDED_TENANT).toBe('common');
  });

  it('accepts clientId and tenantId overrides', () => {
    const app = buildMsalApp({
      cache: new InMemoryTokenCache(),
      clientId: 'other-id',
      tenantId: 'other-tenant',
    });
    expect(app).toBeDefined();
  });

  it('exports the spec scope set', () => {
    expect(OUTLOOK_SCOPES).toEqual([
      'Mail.ReadWrite',
      'Calendars.Read',
      'Calendars.Read.Shared',
      'User.Read',
    ]);
  });

  it('reads cache via the plugin hook on first token-cache access', async () => {
    const cache = new InMemoryTokenCache();
    await cache.save('{"AccessToken":{},"RefreshToken":{},"IdToken":{},"Account":{},"AppMetadata":{}}');
    const app = buildMsalApp({ cache });
    // Forcing a cache read via the public surface:
    const accounts = await app.getTokenCache().getAllAccounts();
    expect(accounts).toEqual([]);
  });
});
