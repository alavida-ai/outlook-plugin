import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { homedir } from 'node:os';
import { join } from 'node:path';

import {
  _resetClientForTesting,
  getClient,
  resolveCachePath,
  type PluginConfig,
} from './client.js';

describe('resolveCachePath — 4-level precedence', () => {
  beforeEach(() => {
    delete process.env.OUTLOOK_TOKEN_CACHE;
  });
  afterEach(() => {
    delete process.env.OUTLOOK_TOKEN_CACHE;
  });

  it('explicit tokenCachePath wins over everything', () => {
    process.env.OUTLOOK_TOKEN_CACHE = '/env/path/tokens.json';
    const path = resolveCachePath({
      tokenCachePath: '/explicit/path/tokens.json',
      agentDir: '/agent/dir',
    });
    expect(path).toBe('/explicit/path/tokens.json');
  });

  it('OUTLOOK_TOKEN_CACHE env wins over agentDir', () => {
    process.env.OUTLOOK_TOKEN_CACHE = '/env/path/tokens.json';
    const path = resolveCachePath({ agentDir: '/openclaw/agents/alfred/agent' });
    expect(path).toBe('/env/path/tokens.json');
  });

  it('agentDir resolves to <agentDir>/outlook-tokens.json', () => {
    const path = resolveCachePath({ agentDir: '/openclaw/agents/alfred/agent' });
    expect(path).toBe('/openclaw/agents/alfred/agent/outlook-tokens.json');
  });

  it('falls back to ~/.outlook-plugin/tokens.json when nothing is provided', () => {
    const path = resolveCachePath({});
    expect(path).toBe(join(homedir(), '.outlook-plugin', 'tokens.json'));
  });
});

describe('getClient — per-agent memoisation', () => {
  beforeEach(() => {
    _resetClientForTesting();
    delete process.env.OUTLOOK_TOKEN_CACHE;
  });
  afterEach(() => {
    _resetClientForTesting();
    delete process.env.OUTLOOK_TOKEN_CACHE;
  });

  it('returns the same client for the same resolved cache path', () => {
    const a = getClient({ agentDir: '/agents/alfred/agent' });
    const b = getClient({ agentDir: '/agents/alfred/agent' });
    expect(a).toBe(b);
  });

  it('returns DIFFERENT clients for different agentDir values', () => {
    const a = getClient({ agentDir: '/agents/alfred/agent' });
    const b = getClient({ agentDir: '/agents/baerbel/agent' });
    expect(a).not.toBe(b);
  });

  it('returns different clients when clientId differs (operator override)', () => {
    const a = getClient({ agentDir: '/agents/a/agent', clientId: 'app-1' });
    const b = getClient({ agentDir: '/agents/a/agent', clientId: 'app-2' });
    expect(a).not.toBe(b);
  });

  it('memoisation key includes resolved path — explicit override and matching default collapse to one entry', () => {
    // If two callers ask for the same agentDir, but one passes an explicit
    // tokenCachePath that happens to point at the same file, they should
    // collapse to one client.
    const explicit = '/agents/alfred/agent/outlook-tokens.json';
    const a = getClient({ tokenCachePath: explicit });
    const b = getClient({ agentDir: '/agents/alfred/agent' });
    expect(a).toBe(b);
  });

  it('_resetClientForTesting clears the cache', () => {
    const a = getClient({ agentDir: '/agents/alfred/agent' });
    _resetClientForTesting();
    const b = getClient({ agentDir: '/agents/alfred/agent' });
    expect(a).not.toBe(b);
  });

  it('accepts an empty config and returns a standalone-path client', () => {
    const a = getClient({});
    const b = getClient({});
    expect(a).toBe(b);
  });
});

describe('getClient — config surface', () => {
  beforeEach(() => _resetClientForTesting());
  afterEach(() => _resetClientForTesting());

  it('passes account.upn through to graph auth provider (no throw)', () => {
    // Smoke test — just ensure construction succeeds with a UPN pin.
    const c = getClient({ account: 'alice@example.com', agentDir: '/agents/a/agent' });
    expect(c).toBeDefined();
  });

  it('accepts agentId without breaking (currently unused in cache path, but on PluginConfig)', () => {
    const cfg: PluginConfig = { agentId: 'alfred', agentDir: '/agents/alfred/agent' };
    const c = getClient(cfg);
    expect(c).toBeDefined();
  });
});
