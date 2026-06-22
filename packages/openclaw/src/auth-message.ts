/**
 * Out-of-band delivery of the sign-in URL.
 *
 * `outlook_auth_login` must not return the auth URL to the agent: a
 * prompt-injected agent could swap it for a phishing link, and the user — primed
 * to expect a "sign in to Microsoft" link — would likely trust it. Instead the
 * tool **stashes** the URL here (keyed by `sessionKey`) and returns the agent
 * only a sanitized acknowledgement. A `message_sending` hook then rewrites the
 * agent's next outbound reply in that session to carry the canonical link.
 *
 * The agent therefore never possesses the URL, so it cannot alter it; the
 * delivered link is always the verbatim, server-generated one.
 *
 * Note: external plugins cannot originate a standalone channel message
 * (`sendSessionAttachment`/`scheduleSessionTurn` are bundled-only), so we ride
 * the agent's own outbound reply via `message_sending`, which is rewrite-only.
 */

/**
 * How long a stashed URL waits for the agent's reply to carry it. The reply
 * normally lands within seconds; this bounds how long a link could otherwise
 * attach to an unrelated later message in the same session.
 */
export const AUTH_MESSAGE_TTL_MS = 5 * 60 * 1000;

interface PendingAuthMessage {
  url: string;
  expiresAt: number;
}

// Module-level: shared between the tool (writer) and the hook (reader). One
// entry per in-flight sign-in per session.
const pendingAuthMessages = new Map<string, PendingAuthMessage>();

/** Stash the sign-in URL for `sessionKey` until `expiresAt` (epoch ms). */
export function stashAuthMessage(sessionKey: string, url: string, expiresAt: number): void {
  pendingAuthMessages.set(sessionKey, { url, expiresAt });
}

/**
 * Single-use take: returns the stashed URL for `sessionKey` and removes it, or
 * `null` if none is pending or it has expired (expired entries are dropped).
 */
export function takeAuthMessage(sessionKey: string, now: number = Date.now()): string | null {
  const entry = pendingAuthMessages.get(sessionKey);
  if (!entry) return null;
  pendingAuthMessages.delete(sessionKey);
  if (now > entry.expiresAt) return null;
  return entry.url;
}

/** Test/diagnostic helper: number of pending messages held. */
export function pendingAuthMessageCount(): number {
  return pendingAuthMessages.size;
}

/** Test helper: drop all pending messages. */
export function clearAuthMessages(): void {
  pendingAuthMessages.clear();
}

/** Remove expired entries. Safe to call often. */
export function gcExpiredAuthMessages(now: number = Date.now()): void {
  for (const [key, entry] of pendingAuthMessages) {
    if (now > entry.expiresAt) pendingAuthMessages.delete(key);
  }
}

/** Render the canonical, self-contained sign-in message delivered to the user. */
export function renderAuthMessage(url: string): string {
  return (
    'Click the link below to sign in to Microsoft Outlook:\n\n' +
    `${url}\n\n` +
    "Once you've signed in, let me know and I'll confirm the connection."
  );
}

// Minimal structural shapes for the `message_sending` hook. The full SDK types
// (PluginHookMessageSendingEvent / PluginHookMessageContext /
// PluginHookMessageSendingResult) aren't re-exported from the public plugin
// barrel; these subsets are assignable to them at the `api.on(...)` call site.
interface MessageSendingEvent {
  content: string;
}
interface MessageHookContext {
  sessionKey?: string;
}
interface MessageSendingResult {
  content?: string;
}

export type AuthMessageHook = (
  event: MessageSendingEvent,
  ctx: MessageHookContext,
) => MessageSendingResult | void;

/**
 * Build the `message_sending` hook. When the outbound message's session has a
 * pending sign-in URL, it **replaces** the content with the canonical link
 * message (replace, not append, so the agent can't wrap the link in misleading
 * text) and consumes the entry. Otherwise it returns nothing (no decision).
 */
export function makeAuthMessageHook(now: () => number = () => Date.now()): AuthMessageHook {
  return (_event, ctx) => {
    if (!ctx.sessionKey) return;
    const url = takeAuthMessage(ctx.sessionKey, now());
    if (!url) return;
    return { content: renderAuthMessage(url) };
  };
}
