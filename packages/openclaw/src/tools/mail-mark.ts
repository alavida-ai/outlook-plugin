/**
 * `mail_mark` — set a message's read state.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  messageId: Type.String({ description: 'Message id to mark.' }),
  state: Type.String({
    enum: ['read', 'unread'] as const,
    description: "Set to 'read' or 'unread'.",
  }),
});

const mailMark = defineTool({
  name: 'mail_mark',
  description: "Mark a message read or unread.",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const isRead = params.state === 'read';
    return client.mail.mark(params.messageId, isRead);
  },
});

export default mailMark;
