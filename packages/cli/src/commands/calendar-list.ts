import { parseArgs } from 'node:util';

import type { EventSummary } from '@alavida-ai/outlook-core';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook calendar list [options]

List events in a date range. Uses /me/calendarView — recurring events are
expanded into individual occurrences.

Options:
  -d, --days N         How many days forward (default: 7). Ignored if --after
                       and --before are both supplied.
      --after ISO      Window start (YYYY-MM-DD or full ISO 8601). Default: now.
      --before ISO     Window end. Default: now + --days days.
  -n, --limit N        Max events (default: 50).
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        days: { type: 'string', short: 'd' },
        after: { type: 'string' },
        before: { type: 'string' },
        limit: { type: 'string', short: 'n' },
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

  let limit = 50;
  if (parsed.values.limit !== undefined) {
    const n = Number.parseInt(parsed.values.limit, 10);
    if (!Number.isFinite(n) || n <= 0) {
      eprintln(`Invalid --limit: ${parsed.values.limit}`);
      return 1;
    }
    limit = n;
  }

  let after = parsed.values.after;
  let before = parsed.values.before;
  if (!before && parsed.values.days !== undefined) {
    const d = Number.parseInt(parsed.values.days, 10);
    if (!Number.isFinite(d) || d <= 0) {
      eprintln(`Invalid --days: ${parsed.values.days}`);
      return 1;
    }
    const nowMs = Date.now();
    after = after ?? new Date(nowMs).toISOString();
    before = new Date(nowMs + d * 86_400_000).toISOString();
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const page = await ctx.outlook.calendar.list({ limit, after, before });
    if (parsed.values.json) {
      printJson({
        results: page.results,
        count: page.count,
        nextLink: page.nextLink,
      });
    } else {
      renderEventList(page.results);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function renderEventList(events: EventSummary[]): void {
  if (events.length === 0) {
    println('(no events in window)');
    return;
  }
  println(`Events (${events.length})`);
  for (const e of events) {
    const start = e.start ? e.start.replace('T', ' ').slice(0, 16) : '';
    const subject = e.subject ?? '(no subject)';
    const organizer = e.organizer ? ` (${e.organizer})` : '';
    const location = e.location ? ` [${e.location}]` : '';
    println(`  ${start}  ${subject}${organizer}${location}`);
    if (e.id) println(`    id: ${e.id}`);
  }
}
