import { parseArgs } from 'node:util';

import {
  AVAILABILITY_LEGEND,
  addDaysToDateStr,
  type AvailabilityResult,
  type AvailabilityScheduleSummary,
} from '@alavida-ai/outlook-core';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook calendar availability --emails ADDR [--emails ADDR ...] [options]

Show free/busy schedules. Defaults to the next 7 days from today, rendered
as each person's events grouped by day. Pass --free-slots N to flip into
"find me open windows >= N minutes during working hours" view, or --raw to
see Graph's raw availability digit string.

Window:
      --pivot DATE     Anchor date (YYYY-MM-DD). Default: today in --tz.
  -d, --days N         Days in window (default: 7). Always includes the pivot day.
      --past           Walk backward from pivot. Window becomes [pivot-(days-1), pivot].
      --tz IANA        Timezone for the window + display. Default: system local.

Output:
      (default)        Scheduled events grouped by day.
      --free-slots N   Free windows of >= N minutes during working hours.
      --raw            Graph's raw availabilityView digit string.
                       Legend: ${AVAILABILITY_LEGEND}.
      --interval M     Block size for --raw, in minutes (default: 30).

Other:
      --emails ADDR    Email to check. Repeatable.
      --account UPN    Pick a specific cached account.
      --json           Emit full JSON envelope.
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
        pivot: { type: 'string' },
        past: { type: 'boolean', default: false },
        'free-slots': { type: 'string' },
        raw: { type: 'boolean', default: false },
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
  let freeSlotsMinMinutes: number | null = null;
  if (parsed.values['free-slots'] !== undefined) {
    const n = Number.parseInt(parsed.values['free-slots'], 10);
    if (!Number.isFinite(n) || n <= 0) {
      eprintln(`Invalid --free-slots: ${parsed.values['free-slots']}`);
      return 1;
    }
    freeSlotsMinMinutes = n;
  }
  const tz = parsed.values.tz ?? Intl.DateTimeFormat().resolvedOptions().timeZone;
  const direction: 'asc' | 'desc' = parsed.values.past ? 'desc' : 'asc';

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const result = await ctx.outlook.calendar.availability({
      emails,
      days,
      interval,
      timeZone: tz,
      pivot: parsed.values.pivot,
      direction,
    });
    if (parsed.values.json) {
      printJson(result);
    } else if (parsed.values.raw) {
      renderRaw(result);
    } else if (freeSlotsMinMinutes !== null) {
      renderFreeSlots(result, freeSlotsMinMinutes);
    } else {
      renderSchedules(result);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

// ── Rendering ────────────────────────────────────────────────────────────────

function renderHeader(r: AvailabilityResult, label: string): void {
  // startTime/endTime are wall-clock in r.timeZone (no offset suffix).
  const start = r.startTime.slice(0, 10);
  const end = addDaysToDateStr(r.endTime.slice(0, 10), -1); // end is exclusive
  println(`${label}  ${start} → ${end}  ${r.timeZone}`);
}

/**
 * Default rendering: each person's scheduleItems, grouped by calendar day.
 *
 * Graph already returns datetimes in the chosen TZ (via `Prefer:
 * outlook.timezone="…"`), so we just slice the date and time off the
 * `dateTime` string — no further conversion needed.
 */
function renderSchedules(r: AvailabilityResult): void {
  renderHeader(r, 'Calendar');
  for (const s of r.schedules) {
    println('');
    println(`${s.scheduleId ?? '(unknown)'}`);
    printWorkingHoursLine(s);

    const byDay = groupItemsByDay(s.scheduleItems);
    const dayKeys = [...byDay.keys()].sort();
    if (dayKeys.length === 0) {
      println('  (no events in window)');
      continue;
    }
    for (const day of dayKeys) {
      println(`  ${dayLabel(day)}`);
      for (const item of byDay.get(day) ?? []) {
        const s1 = item.start?.slice(11, 16) ?? '??:??';
        const s2 = item.end?.slice(11, 16) ?? '??:??';
        const subject = item.subject ?? '(no subject)';
        const status = item.status ? ` [${item.status}]` : '';
        const loc = item.location ? ` @ ${item.location}` : '';
        println(`    ${s1}–${s2}  ${subject}${status}${loc}`);
      }
    }
  }
}

/**
 * --free-slots rendering: open windows of >= minMinutes during the
 * person's working hours, computed by subtracting busy scheduleItems
 * from each working day in the window. Free items in scheduleItems are
 * ignored (they don't block a slot).
 */
function renderFreeSlots(r: AvailabilityResult, minMinutes: number): void {
  renderHeader(r, `Free slots ≥ ${minMinutes}m`);
  for (const s of r.schedules) {
    println('');
    println(`${s.scheduleId ?? '(unknown)'}`);
    printWorkingHoursLine(s);

    const wh = s.workingHours;
    if (!wh || !wh.startTime || !wh.endTime) {
      println('  (no working hours available — cannot compute free slots)');
      continue;
    }
    const workStart = (wh.startTime as string).slice(0, 5); // HH:MM
    const workEnd = (wh.endTime as string).slice(0, 5);
    const workDays = new Set(
      (wh.daysOfWeek ?? []).map((d) => String(d).toLowerCase()),
    );

    let dayCount = 0;
    let cursor = r.startTime.slice(0, 10);
    const endExclusive = r.endTime.slice(0, 10);
    while (cursor < endExclusive) {
      if (workDays.has(dayOfWeekName(cursor))) {
        const winStart = `${cursor}T${workStart}:00`;
        const winEnd = `${cursor}T${workEnd}:00`;
        const slots = freeSlotsForDay(s.scheduleItems, winStart, winEnd, minMinutes);
        if (slots.length > 0) {
          dayCount += 1;
          println(`  ${dayLabel(cursor)}`);
          for (const slot of slots) {
            const startHHMM = slot.start.slice(11, 16);
            const endHHMM = slot.end.slice(11, 16);
            println(`    ${startHHMM}–${endHHMM}  (${formatDuration(slot.minutes)})`);
          }
        }
      }
      cursor = addDaysToDateStr(cursor, 1);
    }
    if (dayCount === 0) {
      println('  (no free slots in window)');
    }
  }
}

function renderRaw(r: AvailabilityResult): void {
  renderHeader(r, 'Availability (raw)');
  println(`Legend: ${AVAILABILITY_LEGEND}.`);
  if (r.schedules.length === 0) {
    println('  (no schedules returned)');
    return;
  }
  for (const s of r.schedules) {
    println(`  ${s.scheduleId ?? '(unknown)'}  ${s.availabilityView ?? ''}`);
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function printWorkingHoursLine(s: AvailabilityScheduleSummary): void {
  const wh = s.workingHours;
  if (!wh) return;
  const days = (wh.daysOfWeek ?? []).map(shortDayName).join(' ');
  const start = (wh.startTime ?? '').slice(0, 5);
  const end = (wh.endTime ?? '').slice(0, 5);
  const tzName =
    (wh.timeZone as { name?: string } | null | undefined)?.name ?? '';
  if (days || start || end) {
    println(`  Working hours: ${days}  ${start}–${end}${tzName ? `  ${tzName}` : ''}`);
  }
}

function groupItemsByDay(
  items: AvailabilityScheduleSummary['scheduleItems'],
): Map<string, AvailabilityScheduleSummary['scheduleItems']> {
  const by = new Map<string, AvailabilityScheduleSummary['scheduleItems']>();
  for (const item of items) {
    if (!item.start) continue;
    const day = item.start.slice(0, 10);
    const arr = by.get(day) ?? [];
    arr.push(item);
    by.set(day, arr);
  }
  for (const arr of by.values()) {
    arr.sort((a, b) => (a.start ?? '').localeCompare(b.start ?? ''));
  }
  return by;
}

interface FreeSlot {
  start: string;
  end: string;
  minutes: number;
}

function freeSlotsForDay(
  items: AvailabilityScheduleSummary['scheduleItems'],
  winStart: string,
  winEnd: string,
  minMinutes: number,
): FreeSlot[] {
  // Filter to busy items that overlap this day's working window. "Free"
  // status items don't block a slot; everything else (busy / tentative /
  // OOO / working elsewhere) does.
  const busy = items
    .filter(
      (i) =>
        i.start &&
        i.end &&
        i.status?.toLowerCase() !== 'free' &&
        i.start < winEnd &&
        i.end > winStart,
    )
    .map((i) => ({
      start: i.start! < winStart ? winStart : i.start!,
      end: i.end! > winEnd ? winEnd : i.end!,
    }))
    .sort((a, b) => a.start.localeCompare(b.start));

  // Merge overlaps so adjacent meetings don't fragment a slot.
  const merged: { start: string; end: string }[] = [];
  for (const b of busy) {
    const last = merged[merged.length - 1];
    if (last && b.start <= last.end) {
      if (b.end > last.end) last.end = b.end;
    } else {
      merged.push({ start: b.start, end: b.end });
    }
  }

  const slots: FreeSlot[] = [];
  let cursor = winStart;
  for (const b of merged) {
    if (b.start > cursor) {
      const m = minutesBetween(cursor, b.start);
      if (m >= minMinutes) slots.push({ start: cursor, end: b.start, minutes: m });
    }
    if (b.end > cursor) cursor = b.end;
  }
  if (cursor < winEnd) {
    const m = minutesBetween(cursor, winEnd);
    if (m >= minMinutes) slots.push({ start: cursor, end: winEnd, minutes: m });
  }
  return slots;
}

function minutesBetween(startIso: string, endIso: string): number {
  // startIso / endIso are wall-clock in the chosen TZ. We diff by parsing
  // them as if they were UTC — the offset cancels out so the duration is
  // correct. (Whole-day DST jumps don't apply because both ends are in
  // the same day.)
  const a = Date.parse(`${startIso}Z`);
  const b = Date.parse(`${endIso}Z`);
  return Math.round((b - a) / 60_000);
}

function formatDuration(minutes: number): string {
  if (minutes < 60) return `${minutes}m`;
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return m === 0 ? `${h}h` : `${h}h ${m}m`;
}

const DAY_NAMES = [
  'sunday',
  'monday',
  'tuesday',
  'wednesday',
  'thursday',
  'friday',
  'saturday',
] as const;

function dayOfWeekName(dateStr: string): string {
  // Parse as UTC to avoid local-tz drift; for date-only math this is exact.
  const d = new Date(`${dateStr}T00:00:00Z`);
  return DAY_NAMES[d.getUTCDay()] ?? '';
}

function shortDayName(d: string): string {
  const s = d.slice(0, 3);
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function dayLabel(dateStr: string): string {
  const d = new Date(`${dateStr}T00:00:00Z`);
  const short = DAY_NAMES[d.getUTCDay()] ?? '';
  return `${short.charAt(0).toUpperCase()}${short.slice(1, 3)} ${dateStr}`;
}
