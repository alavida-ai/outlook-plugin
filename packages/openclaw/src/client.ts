import { homedir } from 'node:os';
import { join } from 'node:path';
import {
  buildMsalApp,
  FileTokenCache,
  makeGraphClient,
  OutlookClient,
  type TokenCache,
} from '@alavida-ai/outlook-core';

export interface PluginConfig {
  clientId?: string;
  tenantId?: string;
  tokenCachePath?: string;
  account?: string;
}

function defaultCachePath(): string {
  return process.env.OUTLOOK_TOKEN_CACHE ?? join(homedir(), '.outlook-plugin', 'tokens.json');
}

let cachedClient: { config: PluginConfig; client: OutlookClient } | null = null;

/**
 * Build (or reuse) the OutlookClient for the supplied plugin config.
 *
 * The plugin SDK calls our tools many times per session; constructing the
 * MSAL app and FileTokenCache on every call is wasteful. We memoise on a
 * structural-equality check of the config; if the operator hot-reloads the
 * config with new values, the next call rebuilds.
 */
export function getClient(config: PluginConfig): OutlookClient {
  if (cachedClient && shallowEqualConfig(cachedClient.config, config)) {
    return cachedClient.client;
  }
  const cache: TokenCache = new FileTokenCache(config.tokenCachePath ?? defaultCachePath());
  const app = buildMsalApp({
    cache,
    clientId: config.clientId,
    tenantId: config.tenantId,
  });
  const graph = makeGraphClient({ app, cache, preferredUpn: config.account });
  const client = new OutlookClient(graph);
  cachedClient = { config: { ...config }, client };
  return client;
}

/** Test-only: reset the memoised client. Exported so unit tests can isolate. */
export function _resetClientForTesting(): void {
  cachedClient = null;
}

function shallowEqualConfig(a: PluginConfig, b: PluginConfig): boolean {
  return (
    a.clientId === b.clientId &&
    a.tenantId === b.tenantId &&
    a.tokenCachePath === b.tokenCachePath &&
    a.account === b.account
  );
}
