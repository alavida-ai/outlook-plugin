/**
 * The auth sign-in URL must reach the user verbatim WITHOUT passing through the
 * agent (so prompt injection can't swap in a phishing link). The tool stashes
 * the URL keyed by sessionKey; a `message_sending` hook then rewrites the
 * agent's next outbound reply in that session to carry the canonical link.
 *
 * These tests cover the server-side map (single-use, TTL) and the hook's
 * rewrite behaviour in isolation.
 */
import { afterEach, describe, expect, it } from 'vitest';

import {
  AUTH_MESSAGE_TTL_MS,
  stashAuthMessage,
  takeAuthMessage,
  pendingAuthMessageCount,
  clearAuthMessages,
  gcExpiredAuthMessages,
  makeAuthMessageHook,
} from './auth-message.js';

const URL = 'https://login.microsoftonline.com/tenant/oauth2/v2.0/authorize?x=1';

afterEach(() => clearAuthMessages());

describe('pending auth-message map', () => {
  it('stash then take returns the URL', () => {
    stashAuthMessage('sess-1', URL, Date.now() + AUTH_MESSAGE_TTL_MS);
    expect(takeAuthMessage('sess-1')).toBe(URL);
  });

  it('take is single-use — a second take returns null', () => {
    stashAuthMessage('sess-1', URL, Date.now() + AUTH_MESSAGE_TTL_MS);
    expect(takeAuthMessage('sess-1')).toBe(URL);
    expect(takeAuthMessage('sess-1')).toBeNull();
    expect(pendingAuthMessageCount()).toBe(0);
  });

  it('take returns null when nothing is stashed for the session', () => {
    expect(takeAuthMessage('unknown')).toBeNull();
  });

  it('take returns null and drops the entry when expired', () => {
    stashAuthMessage('sess-1', URL, Date.now() - 1);
    expect(takeAuthMessage('sess-1')).toBeNull();
    expect(pendingAuthMessageCount()).toBe(0);
  });

  it('gcExpiredAuthMessages drops only expired entries', () => {
    stashAuthMessage('live', URL, Date.now() + AUTH_MESSAGE_TTL_MS);
    stashAuthMessage('dead', URL, Date.now() - 1);
    gcExpiredAuthMessages();
    expect(pendingAuthMessageCount()).toBe(1);
    expect(takeAuthMessage('live')).toBe(URL);
  });
});

describe('makeAuthMessageHook — message_sending rewrite', () => {
  function ctx(sessionKey?: string) {
    return { channelId: 'c', sessionKey } as { channelId: string; sessionKey?: string };
  }
  const event = { to: 'user', content: 'I have sent you the sign-in link.' };

  it('rewrites outbound content to the verbatim URL when the session has a pending link', () => {
    const hook = makeAuthMessageHook();
    stashAuthMessage('sess-1', URL, Date.now() + AUTH_MESSAGE_TTL_MS);

    const result = hook(event, ctx('sess-1'));

    expect(result).toBeTruthy();
    expect(result!.content).toContain(URL);
    // The agent's original wording is replaced, not appended, so it can't wrap
    // the link in misleading text.
    expect(result!.content).not.toContain('I have sent you');
    // Consumed — single delivery.
    expect(pendingAuthMessageCount()).toBe(0);
  });

  it('does nothing (no rewrite) when the session has no pending link', () => {
    const hook = makeAuthMessageHook();
    expect(hook(event, ctx('sess-1'))).toBeUndefined();
  });

  it('does nothing when the context carries no sessionKey', () => {
    const hook = makeAuthMessageHook();
    stashAuthMessage('sess-1', URL, Date.now() + AUTH_MESSAGE_TTL_MS);
    expect(hook(event, ctx(undefined))).toBeUndefined();
    // The pending link is left intact for its real session.
    expect(pendingAuthMessageCount()).toBe(1);
  });

  it('does not rewrite with an expired link', () => {
    const hook = makeAuthMessageHook();
    stashAuthMessage('sess-1', URL, Date.now() - 1);
    expect(hook(event, ctx('sess-1'))).toBeUndefined();
  });
});
