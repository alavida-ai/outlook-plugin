import { GraphError } from '@microsoft/microsoft-graph-client';

import { AuthError, AuthRefreshFailedError } from '../auth/errors.js';

/** Base for every non-auth runtime error from `core`. */
export class CoreError extends Error {
  constructor(message: string) {
    super(message);
    this.name = new.target.name;
  }
}

export class NotFoundError extends CoreError {}
export class ServerError extends CoreError {}
export class NetworkError extends CoreError {}

export class ThrottledError extends CoreError {
  constructor(message: string, public readonly retryAfterSeconds: number | null) {
    super(message);
  }
}

/**
 * Thrown when a mail operation references a message inside the agent's
 * quarantine window — fresh inbound mail the agent is not allowed to read,
 * draft from, or otherwise expose to the model. Threat model: one-time
 * passwords / security codes the user wants to keep out of the agent's
 * context. See `MAIL_QUARANTINE_MINUTES` in `resources/mail.ts`.
 */
export class MailQuarantinedError extends CoreError {
  constructor(
    message: string,
    public readonly messageId: string,
    /** When the message arrived (ISO 8601, UTC). */
    public readonly receivedDateTime: string,
    /** Earliest UTC ISO timestamp at which the agent may read it. */
    public readonly availableAt: string,
  ) {
    super(message);
  }
}

/**
 * Lift any error thrown by the Graph SDK into our taxonomy. Callers wrap
 * every `client.api(...).get/post/etc()` with this so consumers never see a
 * bare `GraphError`.
 */
export function liftGraphError(err: unknown): Error {
  if (err instanceof GraphError) {
    const status = err.statusCode;
    if (status === 401) return new AuthRefreshFailedError(err.message || 'unauthorised');
    if (status === 404) return new NotFoundError(err.message || 'not found');
    if (status === 429) {
      const headers = (err as unknown as { headers?: Record<string, string> }).headers ?? {};
      const ra = parseInt(headers['retry-after'] ?? '', 10);
      return new ThrottledError(err.message || 'throttled', Number.isFinite(ra) ? ra : null);
    }
    if (status >= 500) return new ServerError(err.message || `graph ${status}`);
    return new CoreError(`graph ${status}: ${err.message ?? ''}`.trim());
  }
  if (err instanceof AuthError) return err;
  if (err instanceof Error) return new NetworkError(err.message || 'network error');
  return new NetworkError(String(err));
}
