/**
 * `mail_forward` — draft a forward of an existing message.
 *
 * Graph's createForward composes the forward draft from the original
 * message and prepends the supplied `comment` above the quoted body.
 * Never sends.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  messageId: Type.String({
    description: 'Graph message id of the message being forwarded.',
  }),
  to: Type.Array(Type.String(), {
    description: 'Recipient email addresses.',
    minItems: 1,
  }),
  cc: Type.Optional(
    Type.Array(Type.String(), { description: 'CC email addresses.' }),
  ),
  comment: Type.Optional(
    Type.String({
      description:
        'Optional note prepended above the quoted original. Plain text.',
    }),
  ),
});

const mailForward = defineTool({
  name: 'outlook_mail_forward',
  description:
    "Create a draft forward of a message. Never sends. Graph prepends `comment` above the original; the draft lives in Drafts for the human to review. Returns the draft id and a compose URL.",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const summary = await client.mail.forward(params.messageId, {
      to: params.to,
      cc: params.cc,
      comment: params.comment,
    });
    return summary;
  },
});

export default mailForward;
