/**
 * `calendar_respond` — accept / decline / tentatively-accept a meeting invite.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  eventId: Type.String({ description: 'Event id to respond to.' }),
  response: Type.Union(
    [
      Type.Literal('accept'),
      Type.Literal('decline'),
      Type.Literal('tentative'),
    ],
    { description: 'Response to send.' },
  ),
  comment: Type.Optional(
    Type.String({ description: 'Optional note to the organiser.' }),
  ),
  sendResponse: Type.Optional(
    Type.Boolean({
      description: 'Notify the organiser. Defaults to true.',
    }),
  ),
});

const calendarRespond = defineTool({
  name: 'calendar_respond',
  description:
    'Accept, decline, or tentatively-accept a meeting invite. Set sendResponse=false to update your status without notifying the organiser.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.calendar.respond(params.eventId, {
      response: params.response,
      comment: params.comment,
      sendResponse: params.sendResponse,
    });
  },
});

export default calendarRespond;
