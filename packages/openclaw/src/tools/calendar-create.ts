/**
 * `calendar_create` — create a calendar event.
 *
 * Sends invites to attendees by default (Graph behaviour on /me/events).
 * Supports the same recurrence presets as the CLI (daily, weekdays, weekly,
 * monthly, yearly) and an optional Teams online-meeting link.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  subject: Type.String({ description: 'Event title.' }),
  start: Type.String({
    description: 'Start (YYYY-MM-DD or ISO 8601, e.g. 2026-04-15T09:00).',
  }),
  end: Type.String({ description: 'End (YYYY-MM-DD or ISO 8601).' }),
  attendees: Type.Optional(
    Type.Array(Type.String(), {
      description: 'Attendee email addresses (all required).',
    }),
  ),
  location: Type.Optional(
    Type.String({ description: 'Location display name.' }),
  ),
  body: Type.Optional(
    Type.String({
      description: 'Event body / description. HTML by default — set bodyContentType to "Text" for plain text.',
    }),
  ),
  bodyContentType: Type.Optional(
    Type.Union([Type.Literal('HTML'), Type.Literal('Text')], {
      description: 'Body content type (default: HTML).',
    }),
  ),
  isAllDay: Type.Optional(
    Type.Boolean({ description: 'Mark the event as all-day.' }),
  ),
  isOnlineMeeting: Type.Optional(
    Type.Boolean({ description: 'Add a Teams meeting link.' }),
  ),
  recurrence: Type.Optional(
    Type.Union(
      [
        Type.Literal('daily'),
        Type.Literal('weekdays'),
        Type.Literal('weekly'),
        Type.Literal('monthly'),
        Type.Literal('yearly'),
      ],
      { description: 'Recurrence preset.' },
    ),
  ),
  timeZone: Type.Optional(
    Type.String({
      description: 'IANA timezone for start/end (default: UTC).',
    }),
  ),
});

const calendarCreate = defineTool({
  name: 'calendar_create',
  description:
    'Create a calendar event. Sends invites to attendees by default. Supports optional Teams online-meeting link and a recurrence preset (daily, weekdays, weekly, monthly, yearly).',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.calendar.create({
      subject: params.subject,
      start: params.start,
      end: params.end,
      attendees: params.attendees,
      location: params.location,
      body: params.body,
      bodyContentType: params.bodyContentType,
      isAllDay: params.isAllDay,
      isOnlineMeeting: params.isOnlineMeeting,
      recurrence: params.recurrence,
      timeZone: params.timeZone,
    });
  },
});

export default calendarCreate;
