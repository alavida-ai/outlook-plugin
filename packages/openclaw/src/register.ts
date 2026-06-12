/**
 * Single registration helper for every outlook tool.
 *
 * Every tool in `src/tools/*.ts` exports a {@link ToolDescriptor}; the plugin
 * entry (`src/index.ts`) loops through them and passes each here. The helper:
 *
 *   1. Injects `output` ('pretty' | 'json') and `help` (boolean) into the
 *      tool's parameter schema.
 *   2. Strips those off `params` before invoking the body.
 *   3. Wraps the body in {@link withErrorMapping}.
 *   4. Routes the result through the pretty renderer or `JSON.stringify`.
 *   5. Short-circuits to an auto-generated manpage when `help: true`.
 */
import {
  Type,
  type TObject,
  type TProperties,
  type TSchema,
  type Static,
} from 'typebox';
import type {
  AnyAgentTool,
  OpenClawPluginApi,
  OpenClawPluginToolContext,
} from 'openclaw/plugin-sdk/plugin-entry';

import type { PluginConfig } from './client.js';
import { isToolErrorEnvelope, withErrorMapping } from './errors.js';
import { renderPretty } from './pretty.js';
import { HelpSchema, OutputModeSchema, type OutputMode } from './shared-schemas.js';

interface ToolResultEnvelope {
  content: Array<{ type: 'text'; text: string }>;
  details: unknown;
  terminate?: boolean;
}

export interface ToolDescriptor<TParameters extends TSchema = TSchema> {
  name: string;
  description: string;
  parameters: TParameters;
  label?: string;
  /** Optional hand-authored manpage; auto-generated from schema if omitted. */
  help?: string;
  execute: (params: Static<TParameters>, config: PluginConfig) => Promise<unknown>;
}

export function readPluginConfig(api: OpenClawPluginApi): PluginConfig {
  const raw = (api.pluginConfig ?? {}) as Record<string, unknown>;
  return {
    clientId: typeof raw.clientId === 'string' ? raw.clientId : undefined,
    tenantId: typeof raw.tenantId === 'string' ? raw.tenantId : undefined,
    tokenCachePath: typeof raw.tokenCachePath === 'string' ? raw.tokenCachePath : undefined,
    account: typeof raw.account === 'string' ? raw.account : undefined,
    oauthRedirectUri:
      typeof raw.oauthRedirectUri === 'string' ? raw.oauthRedirectUri : undefined,
  };
}

function withSharedParams<TParameters extends TSchema>(parameters: TParameters): TSchema {
  const asObject = parameters as unknown as TObject & { type?: string };
  if (!asObject || asObject.type !== 'object' || !asObject.properties) {
    return parameters;
  }
  const augmented: TProperties = {
    ...(asObject.properties as TProperties),
    output: OutputModeSchema,
    help: HelpSchema,
  };
  return Type.Object(augmented) as unknown as TSchema;
}

function splitMeta(params: unknown): {
  outputMode: OutputMode;
  helpRequested: boolean;
  toolParams: unknown;
} {
  if (typeof params !== 'object' || params === null) {
    return { outputMode: 'pretty', helpRequested: false, toolParams: params };
  }
  const { output, help, ...rest } = params as Record<string, unknown>;
  return {
    outputMode: output === 'json' ? 'json' : 'pretty',
    helpRequested: help === true,
    toolParams: rest,
  };
}

function buildHelpText<P extends TSchema>(reg: ToolDescriptor<P>): string {
  if (reg.help) return reg.help;

  const params = reg.parameters as unknown as {
    properties?: Record<string, { description?: string; type?: string }>;
    required?: string[];
  };
  const required = new Set(params.required ?? []);
  const headline = reg.description.split(/(?<=[.!?])\s/)[0] ?? reg.description;

  const lines: string[] = [
    'NAME',
    `  ${reg.name} — ${headline}`,
    '',
    'DESCRIPTION',
    `  ${reg.description}`,
    '',
    'PARAMETERS',
  ];

  const props = params.properties ?? {};
  if (Object.keys(props).length === 0) {
    lines.push('  (none)');
  } else {
    for (const [key, schema] of Object.entries(props)) {
      const flag = required.has(key) ? 'required' : 'optional';
      lines.push(`  ${key.padEnd(14)} (${flag}) ${schema.description ?? ''}`);
    }
  }

  lines.push(
    '',
    'OUTPUT',
    "  output: 'pretty' (default) for a summary; 'json' for the raw payload.",
    '  help: true   redisplay this page.',
  );
  return lines.join('\n');
}

function toResult(payload: unknown, mode: OutputMode): ToolResultEnvelope {
  const text = mode === 'json' ? safeStringify(payload) : renderPretty(payload);
  return { content: [{ type: 'text', text }], details: payload };
}

function safeStringify(value: unknown): string {
  if (typeof value === 'string') return value;
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

export function registerTool<TParameters extends TSchema>(
  api: OpenClawPluginApi,
  descriptor: ToolDescriptor<TParameters>,
  getConfig: () => PluginConfig,
): void {
  // Pass a FACTORY (not a static tool) so the host invokes us once per
  // agent setup with that agent's trusted context (`agentId`, `agentDir`,
  // `sessionKey`, …). The factory closes over the per-agent ctx and bakes
  // it into the per-execute config the tool body receives.
  // Second arg `{ name }` is the inspector hint — factories are opaque at
  // inspect time, so the SDK needs it upfront, else tools show as `(anonymous)`.
  api.registerTool((ctx: OpenClawPluginToolContext): AnyAgentTool => ({
    name: descriptor.name,
    description: descriptor.description,
    parameters: withSharedParams(descriptor.parameters),
    label: descriptor.label ?? descriptor.name,
    async execute(_toolCallId: string, params: unknown): Promise<ToolResultEnvelope> {
      const meta = splitMeta(params);
      if (meta.helpRequested) {
        const helpText = buildHelpText(descriptor);
        return {
          content: [{ type: 'text', text: helpText }],
          details: { help: helpText, tool: descriptor.name },
        };
      }
      // Read plugin config fresh each call (supports hot-reload), overlay
      // with the agent-scoped context captured at factory time.
      const config: PluginConfig = {
        ...getConfig(),
        agentId: ctx.agentId,
        agentDir: ctx.agentDir,
      };
      const result = await withErrorMapping(descriptor.name, () =>
        descriptor.execute(meta.toolParams as Static<TParameters>, config),
      );
      return toResult(result, meta.outputMode);
    },
  }), { name: descriptor.name });
}

export function defineTool<TParameters extends TSchema>(
  descriptor: ToolDescriptor<TParameters>,
): ToolDescriptor<TParameters> {
  return descriptor;
}

export { isToolErrorEnvelope };
