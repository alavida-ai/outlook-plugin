/**
 * Error-to-tool-response mapping for the outlook plugin.
 *
 * Every typed error from `@alavida-ai/outlook-core` is mapped to a stable,
 * machine-readable response the agent can branch on. Tool execute bodies
 * wrap themselves in {@link withErrorMapping} so OpenClaw never sees a raw
 * exception — the agent always gets either the normal return value or a
 * `{ __toolError: {...} }` envelope.
 */
import {
  AuthAmbiguousAccountError,
  AuthCacheCorruptError,
  AuthCacheMissingError,
  AuthError,
  AuthInteractionRequiredError,
  AuthLockTimeoutError,
  AuthRefreshFailedError,
  CoreError,
  MailQuarantinedError,
  NetworkError,
  NotFoundError,
  ServerError,
  ThrottledError,
} from '@alavida-ai/outlook-core';

export interface ToolErrorResponse {
  /** Machine-readable error code (stable). */
  error: string;
  /** Human-readable explanation. */
  message: string;
  /** Suggested next step for the agent. Stable across runs. */
  hint?: string;
  /** ThrottledError — `Retry-After` seconds, when present. */
  retryAfterSeconds?: number;
  /** AuthAmbiguousAccountError — cached UPNs the agent can choose from. */
  accounts?: readonly string[];
  /** MailQuarantinedError — earliest UTC ISO timestamp the message can be read. */
  availableAt?: string;
  /** MailQuarantinedError — when the message arrived (UTC ISO). */
  receivedDateTime?: string;
}

export interface ToolErrorEnvelope {
  __toolError: ToolErrorResponse;
}

export function isToolErrorEnvelope(value: unknown): value is ToolErrorEnvelope {
  return (
    typeof value === 'object' &&
    value !== null &&
    '__toolError' in value &&
    typeof (value as { __toolError: unknown }).__toolError === 'object'
  );
}

/**
 * Wrap a tool body. Returns the body's value on success, or a
 * {@link ToolErrorEnvelope} when any error is thrown.
 */
export async function withErrorMapping<T>(
  toolName: string,
  fn: () => Promise<T>,
): Promise<T | ToolErrorEnvelope> {
  try {
    return await fn();
  } catch (e) {
    return { __toolError: mapError(toolName, e) };
  }
}

function mapError(_toolName: string, e: unknown): ToolErrorResponse {
  if (e instanceof AuthCacheMissingError) {
    return {
      error: 'auth_cache_missing',
      message: e.message,
      hint: e.nextStep,
    };
  }
  if (e instanceof AuthCacheCorruptError) {
    return {
      error: 'auth_cache_corrupt',
      message: e.message,
      hint: e.nextStep,
    };
  }
  if (e instanceof AuthRefreshFailedError) {
    return {
      error: 'auth_refresh_failed',
      message: e.message,
      hint: e.nextStep,
    };
  }
  if (e instanceof AuthInteractionRequiredError) {
    return {
      error: 'auth_interaction_required',
      message: e.message,
      hint: e.nextStep,
    };
  }
  if (e instanceof AuthAmbiguousAccountError) {
    return {
      error: 'auth_ambiguous_account',
      message: e.message,
      hint: e.nextStep,
      accounts: e.accounts,
    };
  }
  if (e instanceof AuthLockTimeoutError) {
    return {
      error: 'auth_lock_timeout',
      message: e.message,
      hint: e.nextStep,
    };
  }
  if (e instanceof AuthError) {
    return { error: 'auth_error', message: e.message, hint: e.nextStep };
  }
  if (e instanceof MailQuarantinedError) {
    return {
      error: 'mail_quarantined',
      message: e.message,
      hint:
        'Inbound messages are hidden for a short window after arrival to keep ' +
        'one-time codes and 2FA out of the agent. Wait until `availableAt` and ' +
        'try again, or ask the user to handle this message themselves.',
      availableAt: e.availableAt,
      receivedDateTime: e.receivedDateTime,
    };
  }
  if (e instanceof NotFoundError) {
    return {
      error: 'not_found',
      message: e.message,
      hint: 'Confirm the id/path with the user. List endpoints will surface valid ids.',
    };
  }
  if (e instanceof ThrottledError) {
    return {
      error: 'rate_limited',
      message: e.message,
      hint: 'Microsoft Graph throttled the request. Wait the suggested duration and retry.',
      retryAfterSeconds: e.retryAfterSeconds ?? undefined,
    };
  }
  if (e instanceof ServerError) {
    return {
      error: 'graph_server_error',
      message: e.message,
      hint: 'Microsoft Graph returned 5xx after retries. Try again later; if persistent, escalate.',
    };
  }
  if (e instanceof NetworkError) {
    return {
      error: 'network_error',
      message: e.message,
      hint: 'Could not reach Microsoft Graph. Check network; the plugin retries transient errors automatically.',
    };
  }
  if (e instanceof CoreError) {
    return { error: 'core_error', message: e.message };
  }
  if (e instanceof Error) {
    return { error: 'unknown_error', message: e.message };
  }
  return { error: 'unknown_error', message: String(e) };
}
