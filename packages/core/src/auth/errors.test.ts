import { describe, expect, it } from 'vitest';
import {
  AuthError,
  AuthCacheMissingError,
  AuthCacheCorruptError,
  AuthRefreshFailedError,
  AuthInteractionRequiredError,
  AuthAmbiguousAccountError,
  AuthLockTimeoutError,
} from './errors.js';

describe('AuthError taxonomy', () => {
  it('every variant is an instance of AuthError', () => {
    expect(new AuthCacheMissingError()).toBeInstanceOf(AuthError);
    expect(new AuthCacheCorruptError('bad json')).toBeInstanceOf(AuthError);
    expect(new AuthRefreshFailedError('refresh rejected')).toBeInstanceOf(AuthError);
    expect(new AuthInteractionRequiredError('MFA required')).toBeInstanceOf(AuthError);
    expect(new AuthAmbiguousAccountError(['a@x.com', 'b@y.com'])).toBeInstanceOf(AuthError);
    expect(new AuthLockTimeoutError(30_000)).toBeInstanceOf(AuthError);
  });

  it('every variant carries a nextStep string', () => {
    const variants: AuthError[] = [
      new AuthCacheMissingError(),
      new AuthCacheCorruptError('x'),
      new AuthRefreshFailedError('x'),
      new AuthInteractionRequiredError('x'),
      new AuthAmbiguousAccountError(['a@x.com']),
      new AuthLockTimeoutError(30_000),
    ];
    for (const e of variants) {
      expect(e.nextStep).toMatch(/outlook auth login|--account|wait/);
    }
  });

  it('AuthAmbiguousAccountError lists UPNs in the message', () => {
    const e = new AuthAmbiguousAccountError(['alice@example.com', 'bob@example.com']);
    expect(e.message).toContain('alice@example.com');
    expect(e.message).toContain('bob@example.com');
    expect(e.accounts).toEqual(['alice@example.com', 'bob@example.com']);
  });
});
