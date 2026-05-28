/**
 * Base class for every auth-time failure raised by `@alavida-ai/outlook-core`.
 *
 * Tools (CLI and OpenClaw) catch `AuthError` and surface `.nextStep` verbatim
 * to the human. Subclasses carry typed context the formatter can use to print
 * something more specific.
 */
export class AuthError extends Error {
  /** Short, user-facing remediation. Always populated. */
  public readonly nextStep: string;

  constructor(message: string, nextStep: string) {
    super(message);
    this.name = new.target.name;
    this.nextStep = nextStep;
  }
}

export class AuthCacheMissingError extends AuthError {
  constructor() {
    super('No cached Microsoft account.', 'Run `outlook auth login` to sign in.');
  }
}

export class AuthCacheCorruptError extends AuthError {
  constructor(reason: string) {
    super(
      `Token cache is unreadable (${reason}).`,
      'Run `outlook auth logout` then `outlook auth login` to start fresh.',
    );
  }
}

export class AuthRefreshFailedError extends AuthError {
  constructor(reason: string) {
    super(
      `Silent token refresh failed: ${reason}.`,
      'Run `outlook auth login` to re-authenticate.',
    );
  }
}

export class AuthInteractionRequiredError extends AuthError {
  constructor(reason: string) {
    super(
      `Microsoft requires interactive sign-in (${reason}).`,
      'Run `outlook auth login` to complete sign-in.',
    );
  }
}

export class AuthAmbiguousAccountError extends AuthError {
  /** UPNs of the cached accounts, in cache order. */
  public readonly accounts: readonly string[];

  constructor(accounts: readonly string[]) {
    super(
      `Multiple accounts cached (${accounts.join(', ')}); none selected.`,
      'Pass `--account <upn>` (CLI), set `OUTLOOK_ACCOUNT=<upn>`, or set the `account` config field (plugin).',
    );
    this.accounts = accounts;
  }
}

export class AuthLockTimeoutError extends AuthError {
  constructor(timeoutMs: number) {
    super(
      `Couldn't acquire token-cache refresh lock within ${timeoutMs} ms.`,
      'Another process is mid-refresh; wait a few seconds and retry, or delete `~/.outlook-cli/tokens.lock` if no other process is using it.',
    );
  }
}
