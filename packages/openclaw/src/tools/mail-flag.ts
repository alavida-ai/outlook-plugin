/**
 * `mail_flag` — set the follow-up flag on a message.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const FLAG_STATES = ['flagged', 'complete', 'notFlagged'] as const;
type FlagState = (typeof FLAG_STATES)[number];

const Params = Type.Object({
  messageId: Type.String({ description: 'Message id to flag.' }),
  state: Type.String({
    enum: FLAG_STATES,
    description: "Follow-up flag status: 'flagged', 'complete', or 'notFlagged'.",
  }),
});

const mailFlag = defineTool({
  name: 'mail_flag',
  description: "Set the follow-up flag on a message ('flagged', 'complete', or 'notFlagged').",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.mail.flag(params.messageId, params.state as FlagState);
  },
});

export default mailFlag;
