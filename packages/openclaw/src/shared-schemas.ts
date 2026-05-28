/**
 * Shared TypeBox schemas for the outlook tool catalogue. Tiny surface for
 * this slice — just the cross-cutting Output and Help params injected by
 * register.ts into every tool.
 *
 * As we add list tools in subsequent slices (mail_list, calendar_list, …)
 * a LimitSchema + PageTokenSchema (or whatever Graph uses) lives here too.
 */
import { Type, type Static } from 'typebox';

/**
 * Shared output-mode toggle injected by `register.ts` into every tool's
 * parameter schema. Tool bodies never see it — the helper strips it before
 * dispatch and routes the result through the pretty or json renderer.
 */
export const OutputModeSchema = Type.Optional(
  Type.String({
    enum: ['pretty', 'json'] as const,
    default: 'pretty',
    description: "'pretty' (default) summary; 'json' raw payload for chaining.",
  }),
);
export type OutputMode = 'pretty' | 'json';

/**
 * Shared help toggle injected by `register.ts`. When `true`, the tool
 * short-circuits to the manpage. Equivalent of a CLI `--help` flag.
 */
export const HelpSchema = Type.Optional(
  Type.Boolean({
    description: 'Return usage docs instead of running. Like <cli> --help.',
  }),
);

/** Re-export Static for tool files. */
export type { Static };
