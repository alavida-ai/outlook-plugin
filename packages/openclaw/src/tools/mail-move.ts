/**
 * `mail_move` — move a message to another folder.
 *
 * Folder accepts a well-known name (inbox, sentitems, drafts, deleteditems,
 * junkemail, archive, outbox, scheduled, clutter), a Graph folder id, or a
 * custom folder displayName which we resolve via /me/mailFolders.
 *
 * Outlook reassigns ids on move; the returned `id` is the new id, `oldId`
 * is the original input.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  messageId: Type.String({ description: 'Message id to move.' }),
  folder: Type.String({
    description:
      'Destination folder: well-known name (inbox, archive, ...), Graph folder id, or custom displayName.',
  }),
});

const mailMove = defineTool({
  name: 'mail_move',
  description:
    "Move a message to another folder. `folder` accepts a well-known name (inbox, sentitems, drafts, deleteditems, junkemail, archive, outbox, scheduled, clutter), a Graph folder id, or a custom folder displayName. Outlook reassigns ids on move; the returned `id` is the new id.",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.mail.move(params.messageId, params.folder);
  },
});

export default mailMove;
