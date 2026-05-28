/**
 * `mail_read` — read a single message in full.
 *
 * Returns subject, sender, recipients, importance, body, webLink. By default
 * the body comes back in HTML; pass `preferText: true` for plain text.
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
    description: 'Graph message id (from mail_list / mail_search).',
  }),
  preferText: Type.Optional(
    Type.Boolean({
      description: 'Request the plain-text body instead of HTML (default: HTML).',
    }),
  ),
});

const mailRead = defineTool({
  name: 'mail_read',
  description:
    'Read a single Outlook message in full (subject, sender, recipients, body, importance, web link). Read-only.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const m = await client.mail.get(params.messageId, { preferText: params.preferText });
    return {
      id: m.id ?? null,
      subject: m.subject ?? null,
      from: m.from?.emailAddress?.address ?? null,
      to: addresses(m.toRecipients),
      cc: addresses(m.ccRecipients),
      bcc: addresses(m.bccRecipients),
      receivedDateTime: m.receivedDateTime ?? null,
      isRead: m.isRead ?? null,
      hasAttachments: m.hasAttachments ?? null,
      importance: m.importance ?? null,
      bodyContentType: m.body?.contentType ?? null,
      body: m.body?.content ?? null,
      webLink: m.webLink ?? null,
    };
  },
});

export default mailRead;
