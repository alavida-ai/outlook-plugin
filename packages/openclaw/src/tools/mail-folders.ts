/**
 * `mail_folders` — list mail folders with unread + total item counts.
 */
import { Type } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const mailFolders = defineTool({
  name: 'outlook_mail_folders',
  description:
    "List the signed-in user's mail folders with unread and total item counts. Read-only.",
  parameters: Type.Object({}),
  async execute(_params, config) {
    const client = getClient(config);
    const page = await client.mail.listFolders();
    return {
      folders: page.results.map((f) => ({
        id: f.id ?? null,
        displayName: f.displayName ?? null,
        unreadItemCount: f.unreadItemCount ?? 0,
        totalItemCount: f.totalItemCount ?? 0,
      })),
      count: page.count,
      nextLink: page.nextLink,
    };
  },
});

export default mailFolders;
