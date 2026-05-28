import { describe, expect, it } from 'vitest';
import type { AccountInfo } from '@azure/msal-node';

import { resolveAccount } from './accounts.js';
import { AuthAmbiguousAccountError, AuthCacheMissingError } from './errors.js';

const a = (username: string): AccountInfo => ({
  homeAccountId: `${username}-home`,
  environment: 'login.microsoftonline.com',
  tenantId: 'tenant-id',
  username,
  localAccountId: `${username}-local`,
});

describe('resolveAccount', () => {
  it('throws AuthCacheMissingError when no accounts cached', () => {
    expect(() => resolveAccount([], undefined)).toThrow(AuthCacheMissingError);
  });

  it('returns the single cached account when there is exactly one', () => {
    const acc = a('only@example.com');
    expect(resolveAccount([acc], undefined)).toBe(acc);
  });

  it('throws AuthAmbiguousAccountError when multiple accounts and no preference', () => {
    expect(() => resolveAccount([a('one@x.com'), a('two@y.com')], undefined)).toThrow(
      AuthAmbiguousAccountError,
    );
  });

  it('selects by UPN when preference is provided', () => {
    const accounts = [a('one@x.com'), a('two@y.com')];
    expect(resolveAccount(accounts, 'two@y.com')).toBe(accounts[1]);
  });

  it('UPN matching is case-insensitive', () => {
    const accounts = [a('Alice@Example.com')];
    expect(resolveAccount(accounts, 'alice@EXAMPLE.com')).toBe(accounts[0]);
  });

  it('throws AuthAmbiguousAccountError listing UPNs when preference does not match', () => {
    const accounts = [a('one@x.com'), a('two@y.com')];
    let caught: unknown;
    try {
      resolveAccount(accounts, 'three@z.com');
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(AuthAmbiguousAccountError);
    const e = caught as AuthAmbiguousAccountError;
    expect(e.accounts).toEqual(['one@x.com', 'two@y.com']);
  });
});
