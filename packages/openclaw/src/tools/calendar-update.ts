/**
 * `calendar_update` — patch an existing event.
 *
 * Only the supplied fields are PATCHed. `start` and `end` must be supplied
 * together if either is touched — Graph rejects partial updates.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  eventId: Type.String({ description: 'Event id to update.' }),
  subject: Type.Optional(Type.String({ description: 'New subject.' })),
  start: Type.Optional(
    Type.String({ description: 'New start (paired with end).' }),
  ),
  end: Type.Optional(
    Type.String({ description: 'New end (paired with start).' }),
  ),
  location: Type.Optional(
    Type.String({ description: 'New location display name.' }),
  ),
  body: Type.Optional(Type.String({ description: 'New body content.' })),
  bodyContentType: Type.Optional(
    Type.Union([Type.Literal('HTML'), Type.Literal('Text')], {
      description: 'Body content type (default: HTML).',
    }),
  ),
  timeZone: Type.Optional(
    Type.String({
      description: 'IANA timezone for start/end (default: UTC).',
    }),
  ),
});

const calendarUpdate = defineTool({
  name: 'calendar_update',
  description:
    'PATCH an existing event with the supplied fields. start and end must be supplied together if either is touched (Graph rejects partial start/end updates).',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.calendar.update(params.eventId, {
      subject: params.subject,
      start: params.start,
      end: params.end,
      location: params.location,
      body: params.body,
      bodyContentType: params.bodyContentType,
      timeZone: params.timeZone,
    });
  },
});

export default calendarUpdate;
