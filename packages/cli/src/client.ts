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
  preferredUpn: string | undefined;
  outlook: OutlookClient;
}

/** Resolve the preferred UPN: --account flag wins, then OUTLOOK_ACCOUNT env. */
export function resolveUpn(accountFlag: string | undefined): string | undefined {
  if (accountFlag) return accountFlag;
  const env = process.env.OUTLOOK_ACCOUNT;
  return env ? env : undefined;
}

/** Resolve the default token-cache path: $OUTLOOK_TOKEN_CACHE > ~/.outlook-cli/tokens.json */
export function defaultCachePath(): string {
  const override = process.env.OUTLOOK_TOKEN_CACHE;
  if (override) return override;
  return join(homedir(), '.outlook-cli', 'tokens.json');
}

export interface MakeContextOptions {
  preferredUpn?: string;
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
  const preferredUpn = opts.preferredUpn;
  const graph = makeGraphClient({ app, cache, preferredUpn });
  return { app, cache, preferredUpn, outlook: new OutlookClient(graph) };
}
