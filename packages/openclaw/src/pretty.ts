/**
 * Shape-detected pretty renderer for outlook tool results.
 *
 * This slice only ships `whoami`, so the renderer just handles the error
 * envelope and falls through to a truncated-JSON dump for everything else.
 * Outlook-specific renderers (whoami profile card, message-list table, etc.)
 * land in subsequent slices alongside the tools that produce those shapes.
 */
import { isToolErrorEnvelope, type ToolErrorEnvelope } from './errors.js';

/** Render an arbitrary tool payload as compact text. */
export function renderPretty(payload: unknown): string {
  if (payload === undefined || payload === null) return '(no result)';

  if (isToolErrorEnvelope(payload)) {
    return renderError(payload);
  }

  // Generic fallback — JSON.stringify (truncated for readability).
  try {
    const text = JSON.stringify(payload, null, 2);
    return text.length > 4000 ? text.slice(0, 4000) + '\n…(truncated; use output: json)' : text;
  } catch {
    return String(payload);
  }
}

function renderError(envelope: ToolErrorEnvelope): string {
  const e = envelope.__toolError;
  const lines = [`✗ ${e.error}`, `  ${e.message}`];
  if (e.hint) lines.push(`  → ${e.hint}`);
  if (e.retryAfterSeconds !== undefined) lines.push(`  retry after: ${e.retryAfterSeconds}s`);
  if (e.accounts && e.accounts.length > 0) {
    lines.push(`  accounts:    ${e.accounts.join(', ')}`);
  }
  return lines.join('\n');
}
