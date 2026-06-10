import { Type } from 'typebox';
import { describe, expect, it } from 'vitest';
import type {
  AnyAgentTool,
  OpenClawPluginApi,
  OpenClawPluginToolContext,
  OpenClawPluginToolFactory,
} from 'openclaw/plugin-sdk/plugin-entry';

import type { PluginConfig } from './client.js';
import { readPluginConfig, registerTool, type ToolDescriptor } from './register.js';

/**
 * Minimal fake api that captures what registerTool hands to api.registerTool.
 * The SDK accepts either a static tool or a factory; we want to assert that
 * we pass a factory (so per-agent isolation works) and exercise it directly.
 */
function makeFakeApi(): {
  api: OpenClawPluginApi;
  lastFactory: () => OpenClawPluginToolFactory;
  pluginConfig: Record<string, unknown>;
} {
  const pluginConfig: Record<string, unknown> = {};
  let captured: AnyAgentTool | OpenClawPluginToolFactory | null = null;
  const api = {
    pluginConfig,
    registerTool(tool: AnyAgentTool | OpenClawPluginToolFactory): void {
      captured = tool;
    },
    // The real api surface is huge — we only need these for the test.
  } as unknown as OpenClawPluginApi;

  return {
    api,
    pluginConfig,
    lastFactory: () => {
      if (typeof captured !== 'function') {
        throw new Error('registerTool did not pass a factory function');
      }
      return captured;
    },
  };
}

function makeDescriptor(
  capture: (config: PluginConfig) => void,
): ToolDescriptor<ReturnType<typeof Type.Object>> {
  return {
    name: 'test_tool',
    description: 'test',
    parameters: Type.Object({}),
    async execute(_params, config) {
      capture(config);
      return { ok: true };
    },
  };
}

async function callExecute(
  factory: OpenClawPluginToolFactory,
  ctx: OpenClawPluginToolContext,
  params: unknown = {},
): Promise<void> {
  const resolved = factory(ctx);
  if (!resolved) throw new Error('factory returned null');
  const tool = Array.isArray(resolved) ? resolved[0] : resolved;
  if (!tool) throw new Error('factory returned empty array');
  // AnyAgentTool.execute signature in our SDK: (toolCallId: string, params: unknown) => Promise<unknown>
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await (tool as any).execute('test-toolcall-1', params);
}

describe('readPluginConfig — oauthRedirectUri', () => {
  it('reads oauthRedirectUri when set to a string', () => {
    const api = {
      pluginConfig: { oauthRedirectUri: 'https://gw.ts.net/outlook/auth-callback' },
    } as unknown as OpenClawPluginApi;
    expect(readPluginConfig(api).oauthRedirectUri).toBe(
      'https://gw.ts.net/outlook/auth-callback',
    );
  });

  it('leaves oauthRedirectUri undefined when absent or non-string', () => {
    const absent = { pluginConfig: {} } as unknown as OpenClawPluginApi;
    expect(readPluginConfig(absent).oauthRedirectUri).toBeUndefined();
    const wrongType = {
      pluginConfig: { oauthRedirectUri: 123 },
    } as unknown as OpenClawPluginApi;
    expect(readPluginConfig(wrongType).oauthRedirectUri).toBeUndefined();
  });
});

describe('registerTool — factory shape + per-agent context capture', () => {
  it('passes a factory function (not a static tool) to api.registerTool', () => {
    const { api, lastFactory } = makeFakeApi();
    const descriptor = makeDescriptor(() => {});
    registerTool(api, descriptor, () => ({}) as PluginConfig);
    expect(typeof lastFactory()).toBe('function');
  });

  it('factory bakes ctx.agentId and ctx.agentDir into the per-execute config', async () => {
    const { api, lastFactory } = makeFakeApi();
    const captured: PluginConfig[] = [];
    const descriptor = makeDescriptor((cfg) => captured.push(cfg));
    registerTool(api, descriptor, () => ({ clientId: 'cid' }) as PluginConfig);

    await callExecute(lastFactory(), {
      toolName: 'test_tool',
      agentId: 'alfred',
      agentDir: '/openclaw/agents/alfred/agent',
    });

    expect(captured).toHaveLength(1);
    expect(captured[0]).toMatchObject({
      clientId: 'cid',
      agentId: 'alfred',
      agentDir: '/openclaw/agents/alfred/agent',
    });
  });

  it('two factory calls for different agents produce independent execute configs', async () => {
    const { api, lastFactory } = makeFakeApi();
    const captured: PluginConfig[] = [];
    const descriptor = makeDescriptor((cfg) => captured.push(cfg));
    registerTool(api, descriptor, () => ({}) as PluginConfig);

    const factory = lastFactory();
    await callExecute(factory, {
      toolName: 'test_tool',
      agentId: 'alfred',
      agentDir: '/agents/alfred/agent',
    });
    await callExecute(factory, {
      toolName: 'test_tool',
      agentId: 'baerbel',
      agentDir: '/agents/baerbel/agent',
    });

    expect(captured[0].agentId).toBe('alfred');
    expect(captured[0].agentDir).toBe('/agents/alfred/agent');
    expect(captured[1].agentId).toBe('baerbel');
    expect(captured[1].agentDir).toBe('/agents/baerbel/agent');
  });

  it('still works when host calls the factory without agent context', async () => {
    const { api, lastFactory } = makeFakeApi();
    const captured: PluginConfig[] = [];
    const descriptor = makeDescriptor((cfg) => captured.push(cfg));
    registerTool(api, descriptor, () => ({ clientId: 'cid' }) as PluginConfig);

    await callExecute(lastFactory(), { toolName: 'test_tool' });

    expect(captured[0]).toMatchObject({ clientId: 'cid' });
    expect(captured[0].agentId).toBeUndefined();
    expect(captured[0].agentDir).toBeUndefined();
  });

  it('reads the plugin config FRESH each execute (hot-reload supported)', async () => {
    const { api, lastFactory } = makeFakeApi();
    const captured: PluginConfig[] = [];
    const descriptor = makeDescriptor((cfg) => captured.push(cfg));
    let pluginConfig: PluginConfig = { clientId: 'first' };
    registerTool(api, descriptor, () => pluginConfig);

    const factory = lastFactory();
    await callExecute(factory, { toolName: 'test_tool', agentId: 'alfred' });

    // Hot-reload: operator changes plugin config between tool calls.
    pluginConfig = { clientId: 'second' };
    await callExecute(factory, { toolName: 'test_tool', agentId: 'alfred' });

    expect(captured[0].clientId).toBe('first');
    expect(captured[1].clientId).toBe('second');
  });

  it('responds to `help: true` with auto-generated manpage without invoking execute', async () => {
    const { api, lastFactory } = makeFakeApi();
    const captured: PluginConfig[] = [];
    const descriptor = makeDescriptor((cfg) => captured.push(cfg));
    registerTool(api, descriptor, () => ({}) as PluginConfig);

    await callExecute(lastFactory(), { toolName: 'test_tool' }, { help: true });

    expect(captured).toHaveLength(0); // execute never ran
  });
});
