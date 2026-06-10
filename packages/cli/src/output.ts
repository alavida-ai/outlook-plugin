import {
  AuthError,
  CoreError,
  MailQuarantinedError,
  NetworkError,
  NotFoundError,
  ServerError,
  ThrottledError,
} from '@alavida-ai/outlook-core';

/** Write a JSON payload to stdout, terminated with a newline. */
export function printJson(value: unknown): void {
  process.stdout.write(JSON.stringify(value, null, 2) + '\n');
}

/** Write a line of text to stderr. */
export function eprintln(line = ''): void {
  process.stderr.write(line + '\n');
}

/** Write a line of text to stdout. */
export function println(line = ''): void {
  process.stdout.write(line + '\n');
}

/**
 * Format any thrown error for stderr. AuthError surfaces .nextStep
 * verbatim; CoreError variants get a one-line summary.
 */
export function formatError(e: unknown): string {
  if (e instanceof AuthError) {
    return `${e.message}\n  Next: ${e.nextStep}`;
  }
  if (e instanceof MailQuarantinedError) {
    return (
      `Mail blocked: ${e.message}\n` +
      `  Received:     ${e.receivedDateTime}\n` +
      `  Available at: ${e.availableAt}`
    );
  }
  if (e instanceof NotFoundError) {
    return `Not found: ${e.message}`;
  }
  if (e instanceof ThrottledError) {
    const ra = e.retryAfterSeconds !== null ? ` Retry after ${e.retryAfterSeconds}s.` : '';
    return `Microsoft Graph throttled the request.${ra}`;
  }
  if (e instanceof ServerError) {
    return `Microsoft Graph server error: ${e.message}`;
  }
  if (e instanceof NetworkError) {
    return `Network error: ${e.message}`;
  }
  if (e instanceof CoreError) {
    return `Outlook core error: ${e.message}`;
  }
  if (e instanceof Error) return `Unexpected error: ${e.message}`;
  return `Unexpected error: ${String(e)}`;
}
