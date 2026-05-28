/**
 * `calendar_delete` — cancel an event.
 *
 * Graph notifies attendees if any. No confirmation prompt — the agent is
 * expected to confirm with the human itself before invoking destructive
 * actions.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  eventId: Type.String({ description: 'Event id to cancel.' }),
});

const calendarDelete = defineTool({
  name: 'calendar_delete',
  description:
    'Cancel a calendar event. Graph notifies attendees if any are present. Irreversible — confirm with the human before calling.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.calendar.delete(params.eventId);
  },
});

export default calendarDelete;
