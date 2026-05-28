/**
 * `mail_list` — list messages in a folder.
 *
 * Mirrors `outlook mail list`. Defaults to the inbox, top 10 newest first.
 * Filter knobs: unread, sender address, after/before dates (YYYY-MM-DD or
 * ISO 8601), focused / other inference classification. Read-only.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  limit: Type.Optional(
    Type.Integer({
      minimum: 1,
      maximum: 200,
      default: 10,
      description: 'Maximum number of messages to return (default 10).',
    }),
  ),
  folder: Type.Optional(
    Type.String({
      description:
        "Folder name (well-known: inbox, sentitems, drafts, deleteditems, archive, ...) or a Graph folder id. Defaults to 'inbox'.",
    }),
  ),
  unread: Type.Optional(
    Type.Boolean({ description: 'Only unread messages.' }),
  ),
  from: Type.Optional(
    Type.String({ description: 'Filter by sender email address (exact match).' }),
  ),
  after: Type.Optional(
    Type.String({
      description: 'Only messages received on/after this date (YYYY-MM-DD or ISO 8601).',
    }),
  ),
  before: Type.Optional(
    Type.String({
      description: 'Only messages received on/before this date (YYYY-MM-DD or ISO 8601).',
    }),
  ),
  focused: Type.Optional(
    Type.Boolean({ description: 'Only Focused Inbox messages.' }),
  ),
  other: Type.Optional(
    Type.Boolean({ description: 'Only Other (non-Focused) messages.' }),
  ),
});

const mailList = defineTool({
  name: 'mail_list',
  description:
    "List messages in a folder (default: inbox). Read-only. Returns id, subject, from, receivedDateTime, isRead, hasAttachments, bodyPreview and webLink for each message — chain id into mail_read for the full body.",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const page = await client.mail.list({
      limit: params.limit,
      folder: params.folder,
      unread: params.unread,
      from: params.from,
      after: params.after,
      before: params.before,
      focused: params.focused,
      other: params.other,
    });
    return {
      messages: page.results.map((m) => ({
        id: m.id ?? null,
        subject: m.subject ?? null,
        from: m.from?.emailAddress?.address ?? null,
        receivedDateTime: m.receivedDateTime ?? null,
        isRead: m.isRead ?? null,
        hasAttachments: m.hasAttachments ?? null,
        bodyPreview: m.bodyPreview ?? null,
        webLink: m.webLink ?? null,
      })),
      count: page.count,
      nextLink: page.nextLink,
    };
  },
});

export default mailList;
