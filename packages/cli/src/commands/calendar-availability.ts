import { parseArgs } from 'node:util';

import {
  AVAILABILITY_LEGEND,
  type AvailabilityResult,
} from '@alavida-ai/outlook-core';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook calendar availability --emails ADDR [--emails ADDR ...] [options]

Check free/busy across one or more users. Returns an availabilityView
string per email (one digit per --interval-minute block).

Legend: ${AVAILABILITY_LEGEND}.

Options:
      --emails ADDR    Email to check. Repeatable.
  -d, --days N         Days forward (default: 7).
      --interval M     Availability-view block size in minutes (default: 30).
      --tz IANA        Timezone for the window (default: UTC).
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        emails: { type: 'string', multiple: true },
        days: { type: 'string', short: 'd' },
        interval: { type: 'string' },
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

  const emails = parsed.values.emails ?? [];
  if (emails.length === 0) {
    eprintln('At least one --emails value is required.');
    return 1;
  }

  let days = 7;
  if (parsed.values.days !== undefined) {
    const d = Number.parseInt(parsed.values.days, 10);
    if (!Number.isFinite(d) || d <= 0) {
      eprintln(`Invalid --days: ${parsed.values.days}`);
      return 1;
    }
    days = d;
  }
  let interval = 30;
  if (parsed.values.interval !== undefined) {
    const m = Number.parseInt(parsed.values.interval, 10);
    if (!Number.isFinite(m) || m <= 0) {
      eprintln(`Invalid --interval: ${parsed.values.interval}`);
      return 1;
    }
    interval = m;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const result = await ctx.outlook.calendar.availability({
      emails,
      days,
      interval,
      timeZone: parsed.values.tz,
    });
    if (parsed.values.json) {
      printJson(result);
    } else {
      renderAvailability(result);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function renderAvailability(r: AvailabilityResult): void {
  println(`Availability  ${r.startTime} -> ${r.endTime}  ${r.timeZone}`);
  println(`Legend: ${AVAILABILITY_LEGEND}`);
  if (r.schedules.length === 0) {
    println('  (no schedules returned)');
    return;
  }
  for (const s of r.schedules) {
    println(`  ${s.scheduleId ?? '(unknown)'}  ${s.availabilityView ?? ''}`);
  }
}
