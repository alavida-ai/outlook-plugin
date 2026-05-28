/**
 * `calendar_availability` — free/busy across one or more users.
 *
 * Calls Graph's getSchedule and returns an `availabilityView` string per
 * email (one digit per `interval`-minute block). Legend: 0=free,
 * 1=tentative, 2=busy, 3=out-of-office, 4=working-elsewhere.
 */
import { Type, type Static } from 'typebox';

import { AVAILABILITY_LEGEND } from '@alavida-ai/outlook-core';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  emails: Type.Array(Type.String(), {
    description: 'Email addresses to check (including yourself).',
    minItems: 1,
  }),
  days: Type.Optional(
    Type.Integer({
      minimum: 1,
      maximum: 60,
      default: 7,
      description: 'Days forward to query (default 7).',
    }),
  ),
  interval: Type.Optional(
    Type.Integer({
      minimum: 5,
      maximum: 1440,
      default: 30,
      description: 'Availability-view block size in minutes (default 30).',
    }),
  ),
  timeZone: Type.Optional(
    Type.String({
      description: 'IANA timezone for the window (default: UTC).',
    }),
  ),
});

const calendarAvailability = defineTool({
  name: 'calendar_availability',
  description: `Check free/busy across one or more users. Returns an availabilityView string per email (one character per interval-minute block). Legend: ${AVAILABILITY_LEGEND}. Read-only.`,
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    return client.calendar.availability({
      emails: params.emails,
      days: params.days,
      interval: params.interval,
      timeZone: params.timeZone,
    });
  },
});

export default calendarAvailability;
