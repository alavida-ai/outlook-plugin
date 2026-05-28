import { parseArgs } from 'node:util';

import type { EventDetail } from '@alavida-ai/outlook-core';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook calendar show <event-id> [options]

Show a single event in full, including the body.

Options:
      --account UPN    Pick a specific cached account.
      --json           Emit full JSON.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
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
  if (!eventId) {
    eprintln('Missing required <event-id>.');
    eprintln(HELP);
    return 1;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const e = await ctx.outlook.calendar.get(eventId);
    if (parsed.values.json) {
      printJson(e);
    } else {
      renderEvent(e);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function renderEvent(e: EventDetail): void {
  println(`Subject: ${e.subject ?? '(no subject)'}`);
  if (e.start) println(`Start:   ${e.start}${e.timeZone ? ` (${e.timeZone})` : ''}`);
  if (e.end) println(`End:     ${e.end}`);
  if (e.location) println(`Where:   ${e.location}`);
  if (e.organizer) println(`Organizer: ${e.organizer}`);
  if (e.attendees.length > 0) {
    println('Attendees:');
    for (const a of e.attendees) {
      const resp = a.response ? ` (${a.response})` : '';
      println(`  - ${a.address ?? '(no address)'}${resp}`);
    }
  }
  if (e.onlineJoinUrl) println(`Join:    ${e.onlineJoinUrl}`);
  println('');
  const body = e.body ?? '';
  if (body.length > 2000) {
    println(body.slice(0, 2000));
    println('…(body truncated; use --json for full text)');
  } else {
    println(body);
  }
  if (e.webLink) {
    println('');
    println(`Open in Outlook: ${e.webLink}`);
  }
}
