/**
 * `mail_list_attachments` — list attachment metadata on a message.
 *
 * Returns id, name, contentType, size, isInline for each attachment. Use
 * `mail_download_attachment` to fetch the bytes for a specific id.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  messageId: Type.String({
    description: 'Graph message id (from mail_list / mail_search).',
  }),
});

const mailListAttachments = defineTool({
  name: 'outlook_mail_list_attachments',
  description:
    'List attachment metadata (id, name, contentType, size, isInline) on an Outlook message. Read-only. Use mail_download_attachment to fetch bytes.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const page = await client.mail.listAttachments(params.messageId);
    return {
      attachments: page.results.map((a) => ({
        id: a.id ?? null,
        name: a.name ?? null,
        contentType: a.contentType ?? null,
        size: a.size ?? null,
        isInline: a.isInline ?? false,
      })),
      count: page.count,
      nextLink: page.nextLink,
    };
  },
});

export default mailListAttachments;
