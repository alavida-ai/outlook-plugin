/**
 * `calendar_show` — full event details for a single event id.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  eventId: Type.String({ description: 'Event id (from calendar_list).' }),
});

const calendarShow = defineTool({
  name: 'calendar_show',
  description:
    'Show a single calendar event in full, including body (HTML), location, attendees and any online-meeting join URL. Read-only.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.calendar.get(params.eventId);
  },
});

export default calendarShow;
