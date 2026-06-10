/**
 * `outlook_calendar_show` — full event details for a single event id.
 *
 * Body defaults to plain text — Teams meeting invites etc. are mostly
 * HTML chrome that wastes context. Set `preferText: false` to get raw HTML.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  eventId: Type.String({ description: 'Event id (from outlook_calendar_list).' }),
  preferText: Type.Optional(
    Type.Boolean({
      default: true,
      description:
        'Return the body as plain text (default: true — Teams invites etc. are mostly HTML chrome). Set false only if you specifically need the raw HTML.',
    }),
  ),
});

const calendarShow = defineTool({
  name: 'outlook_calendar_show',
  description:
    'Show a single calendar event in full, including body, location, attendees and any online-meeting join URL. Read-only. Body is plain text by default (set preferText:false for raw HTML).',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.calendar.get(params.eventId, {
      preferText: params.preferText ?? true,
    });
  },
});

export default calendarShow;
