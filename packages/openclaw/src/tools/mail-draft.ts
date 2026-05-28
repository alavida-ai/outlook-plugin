/**
 * `mail_draft` — create a draft message in the user's Drafts folder.
 *
 * Never sends. Returns the draft id, webLink, and a compose URL the agent
 * can hand to the human to open the draft in edit mode.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  to: Type.Array(Type.String(), {
    description: 'Recipient email addresses.',
    minItems: 1,
  }),
  cc: Type.Optional(
    Type.Array(Type.String(), { description: 'CC email addresses.' }),
  ),
  bcc: Type.Optional(
    Type.Array(Type.String(), { description: 'BCC email addresses.' }),
  ),
  subject: Type.String({ description: 'Email subject.' }),
  body: Type.String({
    description: 'Body content. Plain text unless `html: true`.',
  }),
  html: Type.Optional(
    Type.Boolean({
      description: 'When true, body is sent as HTML; otherwise plain text.',
    }),
  ),
});

const mailDraft = defineTool({
  name: 'mail_draft',
  description:
    "Create a draft message in the user's Drafts folder. Never sends — the human reviews and sends from Outlook. Returns the draft id and a compose URL for one-click open-in-edit-mode.",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const summary = await client.mail.draft({
      subject: params.subject,
      body: params.body,
      html: params.html,
      to: params.to,
      cc: params.cc,
      bcc: params.bcc,
    });
    return summary;
  },
});

export default mailDraft;
