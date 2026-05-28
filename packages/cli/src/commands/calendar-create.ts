import { parseArgs } from 'node:util';

import { RECURRENCE_PRESETS, type RecurrencePreset } from '@alavida-ai/outlook-core';

import { decodeEscapes } from '../escapes.js';
import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook calendar create --subject S --start ISO --end ISO [options]

Create a calendar event. Sends invites to attendees by default.

Options:
      --subject S          Event title (required).
      --start ISO          Start (YYYY-MM-DD or full ISO 8601, required).
      --end ISO            End (required).
      --attendees ADDR     Attendee email. Repeatable.
      --location LOC       Display name of the location.
      --body BODY          Body content. Interprets \\n, \\r, \\t, \\\\.
      --all-day            Mark the event as all-day.
      --online-meeting     Add a Teams meeting link.
      --recurrence PRESET  One of: daily, weekdays, weekly, monthly, yearly.
      --tz IANA            Timezone for start/end (default: UTC).
      --account UPN        Pick a specific cached account.
      --json               Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        subject: { type: 'string' },
        start: { type: 'string' },
        end: { type: 'string' },
        attendees: { type: 'string', multiple: true },
        location: { type: 'string' },
        body: { type: 'string' },
        'all-day': { type: 'boolean', default: false },
        'online-meeting': { type: 'boolean', default: false },
        recurrence: { type: 'string' },
        tz: { type: 'string' },
        account: { type: 'string' },
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
    });
  } catch (err) {
    eprintln(formatError(err));
    eprintln(HELP);
    return 1;
  }

  if (parsed.values.help) {
    process.stdout.write(HELP);
    return 0;
  }

  if (!parsed.values.subject) {
    eprintln('Missing --subject.');
    return 1;
  }
  if (!parsed.values.start) {
    eprintln('Missing --start.');
    return 1;
  }
  if (!parsed.values.end) {
    eprintln('Missing --end.');
    return 1;
  }

  let recurrence: RecurrencePreset | undefined;
  if (parsed.values.recurrence !== undefined) {
    const preset = parsed.values.recurrence as RecurrencePreset;
    if (!RECURRENCE_PRESETS.has(preset)) {
      eprintln(
        `Invalid --recurrence '${parsed.values.recurrence}'. Valid: ${Array.from(RECURRENCE_PRESETS).join(', ')}.`,
      );
      return 1;
    }
    recurrence = preset;
  }

  const body = parsed.values.body !== undefined ? decodeEscapes(parsed.values.body) : undefined;

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const created = await ctx.outlook.calendar.create({
      subject: parsed.values.subject,
      start: parsed.values.start,
      end: parsed.values.end,
      attendees: parsed.values.attendees,
      location: parsed.values.location,
      body,
      isAllDay: parsed.values['all-day'],
      isOnlineMeeting: parsed.values['online-meeting'],
      recurrence,
      timeZone: parsed.values.tz,
    });
    if (parsed.values.json) {
      printJson(created);
    } else {
      println(`Event "${created.subject ?? '(no subject)'}" created.`);
      println(`  id: ${created.id ?? ''}`);
      if (created.start) println(`  start: ${created.start}`);
      if (created.end) println(`  end:   ${created.end}`);
      if (created.onlineJoinUrl) println(`  join:  ${created.onlineJoinUrl}`);
      if (created.webLink) println(`  open in Outlook: ${created.webLink}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
