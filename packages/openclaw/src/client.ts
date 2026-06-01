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
  /**
   * Set by the OpenClaw plugin SDK via the tool-factory context. Used to scope
   * the default cache path to the agent's own state dir.
   */
  agentId?: string;
  /**
   * Set by the OpenClaw plugin SDK via the tool-factory context. When present
   * and no explicit `tokenCachePath` / `OUTLOOK_TOKEN_CACHE` overrides apply,
   * the token cache lives at `<agentDir>/outlook-tokens.json`.
   */
  agentDir?: string;
}

/**
 * Resolve where the token cache for this invocation should live.
 *
 * Precedence (highest first):
 *   1. `config.tokenCachePath` — explicit operator override in plugin config.
 *   2. `OUTLOOK_TOKEN_CACHE` — process-level override (mainly for tests / VPS overrides).
 *   3. `<agentDir>/outlook-tokens.json` — per-agent default when running under
 *      the OpenClaw plugin runtime (`agentDir` comes from the trusted
 *      `OpenClawPluginToolContext`).
 *   4. `~/.outlook-plugin/tokens.json` — standalone fallback (CLI / no host context).
 */
export function resolveCachePath(
  config: Pick<PluginConfig, 'tokenCachePath' | 'agentDir'>,
): string {
  if (config.tokenCachePath) return config.tokenCachePath;
  if (process.env.OUTLOOK_TOKEN_CACHE) return process.env.OUTLOOK_TOKEN_CACHE;
  if (config.agentDir) return join(config.agentDir, 'outlook-tokens.json');
  return join(homedir(), '.outlook-plugin', 'tokens.json');
}

function clientKey(config: PluginConfig): string {
  // Cache path is the primary discriminator. clientId/tenantId folded in
  // because operators can override them and a fresh MSAL app must be minted.
  return [config.clientId ?? '', config.tenantId ?? '', resolveCachePath(config)].join('|');
}

const clientByKey = new Map<string, OutlookClient>();

/**
 * Build (or reuse) the OutlookClient for the supplied plugin config.
 *
 * On a multi-agent OpenClaw gateway the same plugin is invoked by many
 * agents in the same process; each agent gets its own cache file. We
 * memoise per `(clientId, tenantId, resolvedCachePath)` so each agent gets
 * its own MSAL app + token cache instance, and so a single agent's repeat
 * invocations reuse the same instance instead of paying the MSAL init cost
 * on every tool call.
 */
export function getClient(config: PluginConfig): OutlookClient {
  const key = clientKey(config);
  const existing = clientByKey.get(key);
  if (existing) return existing;
  const cache: TokenCache = new FileTokenCache(resolveCachePath(config));
  const app = buildMsalApp({
    cache,
    clientId: config.clientId,
    tenantId: config.tenantId,
  });
  const graph = makeGraphClient({ app, cache, preferredUpn: config.account });
  const client = new OutlookClient(graph);
  clientByKey.set(key, client);
  return client;
}

/** Test-only: clear the memoisation map so tests get fresh clients. */
export function _resetClientForTesting(): void {
  clientByKey.clear();
}
