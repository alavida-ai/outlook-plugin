import { homedir } from 'node:os';
import { join } from 'node:path';
import {
  buildMsalApp,
  FileTokenCache,
  makeGraphClient,
  OutlookClient,
  type TokenCache,
} from '@alavida-ai/outlook-core';

/**
 * Use ReturnType of buildMsalApp so the CLI doesn't take a direct dep on
 * @azure/msal-node — MSAL is a core-level concern.
 */
type MsalApp = ReturnType<typeof buildMsalApp>;

export interface CliContext {
  app: MsalApp;
  cache: TokenCache;
  outlook: OutlookClient;
}

/** Resolve the default token-cache path: $OUTLOOK_TOKEN_CACHE > ~/.outlook-plugin/tokens.json */
export function defaultCachePath(): string {
  return process.env.OUTLOOK_TOKEN_CACHE ?? join(homedir(), '.outlook-plugin', 'tokens.json');
}

export interface MakeContextOptions {
  cachePath?: string;
  clientId?: string;
  tenantId?: string;
}

/**
 * Construct an MSAL app + token cache + Graph client wired together.
 *
 * Every CLI command that hits Graph calls this. Cheap to construct — none
 * of the children touch disk or the network until the first method call.
 */
export function makeContext(opts: MakeContextOptions = {}): CliContext {
  const cache: TokenCache = new FileTokenCache(opts.cachePath ?? defaultCachePath());
  const app = buildMsalApp({
    cache,
    clientId: opts.clientId ?? process.env.AZURE_CLIENT_ID,
    tenantId: opts.tenantId ?? process.env.AZURE_TENANT_ID,
  });
  // The CLI is single-account (sign-in replaces the cached account), so it
  // never pins a preferred UPN — core resolves the one cached account.
  const graph = makeGraphClient({ app, cache, preferredUpn: undefined });
  return { app, cache, outlook: new OutlookClient(graph) };
}
