import { parseArgs } from 'node:util';

import { ATTENDEE_RESPONSES, type AttendeeResponse } from '@alavida-ai/outlook-core';

import { decodeEscapes } from '../escapes.js';
import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook calendar respond <event-id> <accept|decline|tentative> [options]

Accept, decline, or tentatively-accept a meeting invite. By default the
organiser is notified; pass --no-send to skip the response.

Options:
      --comment TEXT   Optional note to the organiser. Interprets \\n, \\r, \\t, \\\\.
      --send           Notify the organiser (default).
      --no-send        Skip notifying the organiser.
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        comment: { type: 'string' },
        send: { type: 'boolean', default: false },
        'no-send': { type: 'boolean', default: false },
        account: { type: 'string' },
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
      allowPositionals: true,
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

  const eventId = parsed.positionals[0];
  const respRaw = parsed.positionals[1];
  if (!eventId || !respRaw) {
    eprintln('Missing required <event-id> and <accept|decline|tentative>.');
    eprintln(HELP);
    return 1;
  }
  const response = respRaw as AttendeeResponse;
  if (!ATTENDEE_RESPONSES.has(response)) {
    eprintln(
      `Invalid response '${respRaw}'. Valid: ${Array.from(ATTENDEE_RESPONSES).join(', ')}.`,
    );
    return 1;
  }

  let sendResponse = true;
  if (parsed.values['no-send']) sendResponse = false;
  // --send wins when supplied alongside --no-send (caller bug); default is true anyway.
  if (parsed.values.send && parsed.values['no-send']) {
    eprintln('--send and --no-send are mutually exclusive.');
    return 1;
  }

  const comment =
    parsed.values.comment !== undefined ? decodeEscapes(parsed.values.comment) : undefined;

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const r = await ctx.outlook.calendar.respond(eventId, {
      response,
      comment,
      sendResponse,
    });
    if (parsed.values.json) {
      printJson(r);
    } else {
      println(`Responded ${r.response} to event ${r.id}.`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
