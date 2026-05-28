import { describe, expect, it } from 'vitest';
import { GraphError } from '@microsoft/microsoft-graph-client';

import {
  CoreError,
  NotFoundError,
  ThrottledError,
  ServerError,
  NetworkError,
  liftGraphError,
} from './errors.js';
import { AuthError } from '../auth/errors.js';

function makeGraphError(statusCode: number, message = 'msg'): GraphError {
  const err = new GraphError(statusCode, message);
  return err;
}

describe('liftGraphError', () => {
  it('maps 401 to AuthError', () => {
    expect(liftGraphError(makeGraphError(401))).toBeInstanceOf(AuthError);
  });

  it('maps 404 to NotFoundError', () => {
    expect(liftGraphError(makeGraphError(404))).toBeInstanceOf(NotFoundError);
  });

  it('maps 429 to ThrottledError with retryAfter when available', () => {
    const err = makeGraphError(429, 'throttled');
    (err as unknown as { headers?: Record<string, string> }).headers = { 'retry-after': '17' };
    const lifted = liftGraphError(err) as ThrottledError;
    expect(lifted).toBeInstanceOf(ThrottledError);
    expect(lifted.retryAfterSeconds).toBe(17);
  });

  it('maps 503 to ServerError', () => {
    expect(liftGraphError(makeGraphError(503))).toBeInstanceOf(ServerError);
  });

  it('wraps unknown thrown values as NetworkError', () => {
    expect(liftGraphError(new TypeError('fetch failed'))).toBeInstanceOf(NetworkError);
  });

  it('passes through non-Error thrown values as NetworkError', () => {
    expect(liftGraphError('weird')).toBeInstanceOf(NetworkError);
  });

  it('CoreError is the common base', () => {
    expect(liftGraphError(makeGraphError(500))).toBeInstanceOf(CoreError);
  });
});
