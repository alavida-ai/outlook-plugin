/**
 * `outlook_mail_read` — read a single message in full.
 *
 * Returns subject, sender, recipients, importance, body, webLink, inboxLink.
 * Plain-text body is the default (cleaner for LLM consumption); pass
 * `preferText: false` if you specifically need the raw HTML.
 *
 * Refuses inbound non-draft mail younger than 30 minutes (OTP safety window).
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

interface RecipientLike {
  emailAddress?: { address?: string | null } | null;
}

function addresses(list: RecipientLike[] | null | undefined): string[] {
  if (!list) return [];
  const out: string[] = [];
  for (const r of list) {
    const a = r.emailAddress?.address;
    if (a) out.push(a);
  }
  return out;
}

const Params = Type.Object({
  messageId: Type.String({
    description: 'Graph message id (from outlook_mail_list / outlook_mail_search).',
  }),
  preferText: Type.Optional(
    Type.Boolean({
      default: true,
      description:
        'Return the plain-text body instead of HTML (default: true — HTML carries inline styles that are noise for LLM consumption). Set false only when you specifically need the raw HTML.',
    }),
  ),
});

const mailRead = defineTool({
  name: 'outlook_mail_read',
  description:
    'Read a single Outlook message in full (subject, sender, recipients, body, importance, webLink, inboxLink). Read-only. Returns plain text by default; pass preferText:false for raw HTML. Refuses inbound non-draft mail younger than 30 min (safety window for one-time passwords). inboxLink is the outlook.cloud.microsoft URL — share with the user when surfacing the message.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    // Default to plain text — HTML inline styles are mostly noise for an LLM
    // and burn output tokens. preferText explicitly false opts back into HTML.
    const preferText = params.preferText ?? true;
    const m = await client.mail.get(params.messageId, { preferText });
    // Resolve the cloud.microsoft inbox URL via one extra round trip (cheap).
    let inboxLink: string | null = null;
    if (m.id) {
      try {
        const links = await client.mail.inboxLinks([m.id]);
        inboxLink = links[m.id] ?? null;
      } catch {
        // Translation failures are non-fatal — leave inboxLink null.
      }
    }
    return {
      id: m.id ?? null,
      subject: m.subject ?? null,
      from: m.from?.emailAddress?.address ?? null,
      to: addresses(m.toRecipients),
      cc: addresses(m.ccRecipients),
      bcc: addresses(m.bccRecipients),
      receivedDateTime: m.receivedDateTime ?? null,
      isRead: m.isRead ?? null,
      isDraft: m.isDraft ?? null,
      hasAttachments: m.hasAttachments ?? null,
      importance: m.importance ?? null,
      bodyContentType: m.body?.contentType ?? null,
      body: m.body?.content ?? null,
      webLink: m.webLink ?? null,
      inboxLink,
    };
  },
});

export default mailRead;
