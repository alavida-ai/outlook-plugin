/**
 * `mail_importance` — set the importance level of a message.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const LEVELS = ['low', 'normal', 'high'] as const;
type Level = (typeof LEVELS)[number];

const Params = Type.Object({
  messageId: Type.String({ description: 'Message id to update.' }),
  level: Type.String({
    enum: LEVELS,
    description: "Importance level: 'low', 'normal', or 'high'.",
  }),
});

const mailImportance = defineTool({
  name: 'mail_importance',
  description: "Set the importance level of a message ('low', 'normal', or 'high').",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.mail.importance(params.messageId, params.level as Level);
  },
});

export default mailImportance;
