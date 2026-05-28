import { parseArgs } from 'node:util';

import { decodeEscapes } from '../escapes.js';
import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook calendar update <event-id> [options]

Update an existing event. Only the fields supplied are PATCHed. --start and
--end must be supplied together if either is touched.

Options:
      --subject S      New subject.
      --start ISO      New start (paired with --end).
      --end ISO        New end (paired with --start).
      --location LOC   New display name.
      --body BODY      New body. Interprets \\n, \\r, \\t, \\\\.
      --tz IANA        Timezone for start/end (default: UTC).
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
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
        location: { type: 'string' },
        body: { type: 'string' },
        tz: { type: 'string' },
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

  if (
    (parsed.values.start !== undefined) !==
    (parsed.values.end !== undefined)
  ) {
    eprintln('--start and --end must be supplied together.');
    return 1;
  }

  const body = parsed.values.body !== undefined ? decodeEscapes(parsed.values.body) : undefined;

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const updated = await ctx.outlook.calendar.update(eventId, {
      subject: parsed.values.subject,
      start: parsed.values.start,
      end: parsed.values.end,
      location: parsed.values.location,
      body,
      timeZone: parsed.values.tz,
    });
    if (parsed.values.json) {
      printJson(updated);
    } else {
      println(`Updated ${updated.id ?? eventId}.`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
