/**
 * `mail_reply` — draft a reply (or reply-all) to an existing message.
 *
 * Two-step Graph dance: createReply / createReplyAll seeds the draft with
 * the quoted original, then a PATCH overlays the new body on top. Never
 * sends; the draft lives in Drafts for the human to review.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  messageId: Type.String({
    description: 'Graph message id of the message being replied to.',
  }),
  body: Type.String({
    description: 'Reply body content. Plain text unless `html: true`.',
  }),
  all: Type.Optional(
    Type.Boolean({
      description: 'Reply to everyone on the thread (createReplyAll).',
    }),
  ),
  html: Type.Optional(
    Type.Boolean({
      description: 'When true, body is sent as HTML; otherwise plain text.',
    }),
  ),
});

const mailReply = defineTool({
  name: 'outlook_mail_reply',
  description:
    'Create a draft reply (or reply-all when `all: true`) to a message. Never sends; the draft lives in Drafts for the human to review. Returns the draft id and a compose URL.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const summary = await client.mail.reply(params.messageId, {
      body: params.body,
      html: params.html,
      replyAll: params.all,
    });
    return summary;
  },
});

export default mailReply;
