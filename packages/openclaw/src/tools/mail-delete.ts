/**
 * `mail_delete` — soft-delete a message.
 *
 * Outlook moves the message to Deleted Items rather than hard-deleting.
 * Recoverable until the user empties that folder, so no confirmation
 * prompt is exposed to the agent.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  messageId: Type.String({ description: 'Message id to delete.' }),
});

const mailDelete = defineTool({
  name: 'mail_delete',
  description:
    "Soft-delete a message — Outlook moves it to Deleted Items rather than hard-deleting. Recoverable until the user empties that folder.",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.mail.delete(params.messageId);
  },
});

export default mailDelete;
