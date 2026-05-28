/**
 * `calendar_list` — events in a date range.
 *
 * Mirrors `outlook calendar list`. Uses Graph's calendarView, so recurring
 * events are expanded into individual occurrences. Defaults to the next
 * seven days when neither `after` nor `before` is supplied. Read-only.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  after: Type.Optional(
    Type.String({
      description:
        'Window start (YYYY-MM-DD or full ISO 8601). Defaults to now.',
    }),
  ),
  before: Type.Optional(
    Type.String({
      description:
        'Window end (YYYY-MM-DD or full ISO 8601). Defaults to now + 7 days.',
    }),
  ),
  limit: Type.Optional(
    Type.Integer({
      minimum: 1,
      maximum: 500,
      default: 50,
      description: 'Maximum events to return (default 50).',
    }),
  ),
});

const calendarList = defineTool({
  name: 'calendar_list',
  description:
    'List calendar events in a date range. Uses Graph calendarView — recurring events are expanded into occurrences. Defaults to the next seven days. Read-only.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const page = await client.calendar.list({
      after: params.after,
      before: params.before,
      limit: params.limit,
    });
    return {
      events: page.results,
      count: page.count,
      nextLink: page.nextLink,
    };
  },
});

export default calendarList;
