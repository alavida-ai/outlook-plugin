import type { Client } from '@microsoft/microsoft-graph-client';
import type {
  Attendee,
  DateTimeTimeZone,
  Event,
  Location as GraphLocation,
  ScheduleInformation,
  ScheduleItem,
  WorkingHours,
} from '@microsoft/microsoft-graph-types';

import { liftGraphError } from '../graph/errors.js';

/** Availability view legend: 0=free, 1=tentative, 2=busy, 3=OOO, 4=working-elsewhere. */
export const AVAILABILITY_LEGEND =
  '0=free, 1=tentative, 2=busy, 3=out-of-office, 4=working-elsewhere';

/** Flattened Graph Event used by list endpoints. */
export interface EventSummary {
  id: string | null;
  subject: string | null;
  start: string | null;
  end: string | null;
  timeZone: string | null;
  location: string | null;
  organizer: string | null;
  attendees: AttendeeSummary[];
  isOnlineMeeting: boolean | null;
  onlineJoinUrl: string | null;
  isAllDay: boolean | null;
  isCancelled: boolean | null;
  webLink: string | null;
}

export interface AttendeeSummary {
  address: string | null;
  name: string | null;
  type: string | null;
  response: string | null;
}

/** Single-event detail extends summary with body. */
export interface EventDetail extends EventSummary {
  bodyContentType: string | null;
  body: string | null;
  bodyPreview: string | null;
}

export interface PageEnvelope<T> {
  results: T[];
  count: number;
  nextLink: string | null;
}

export interface GetEventOptions {
  /**
   * If true, send `Prefer: outlook.body-content-type="text"` so Graph
   * returns the body as plain text rather than HTML. Useful for terminal
   * rendering where HTML tags are noise. Default: false (Graph default
   * is HTML).
   */
  preferText?: boolean;
}

export interface ListEventsOptions {
  /** Window start ISO (inclusive). Default: now. */
  after?: string;
  /** Window end ISO (exclusive). Default: now + 7 days. */
  before?: string;
  /** Max events to return. Default: 50. */
  limit?: number;
}

export interface AvailabilityInput {
  emails: string[];
  /**
   * Number of calendar days in the window. Default: 7.
   *
   * Combines with `pivot` and `direction` to define the window:
   *   - `direction: 'asc'`  → `[pivot 00:00, pivot+N days 00:00)`
   *   - `direction: 'desc'` → `[pivot-(N-1) days 00:00, pivot+1 day 00:00)`
   *
   * Window length is always exactly N calendar days, anchored to midnight
   * in the chosen `timeZone` so it lines up with real days.
   */
  days?: number;
  /** Interval in minutes for the availability view. Default: 30. */
  interval?: number;
  /** IANA timezone for the query window. Default: 'UTC'. */
  timeZone?: string;
  /**
   * Anchor date for the window. Accepts `YYYY-MM-DD` (uses midnight in
   * `timeZone`) or a full ISO 8601 datetime (the date portion is taken,
   * time is ignored — the window always starts at midnight). Default:
   * today in `timeZone`.
   */
  pivot?: string;
  /**
   * Direction of the window relative to `pivot`. Default: `'asc'`.
   *
   *   - `'asc'`  → pivot + future. Window is `[pivot, pivot + days)`.
   *   - `'desc'` → past + pivot. Window is `[pivot - (days-1), pivot + 1]`.
   *
   * Both directions always *include* the pivot day itself.
   */
  direction?: 'asc' | 'desc';
}

export interface AvailabilityScheduleSummary {
  scheduleId: string | null;
  availabilityView: string | null;
  scheduleItems: Array<{
    subject: string | null;
    start: string | null;
    end: string | null;
    status: string | null;
    location: string | null;
  }>;
  workingHours: WorkingHours | null;
}

export interface AvailabilityResult {
  emails: string[];
  startTime: string;
  endTime: string;
  timeZone: string;
  interval: number;
  schedules: AvailabilityScheduleSummary[];
}

interface RawPageResponse<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

const EVENT_SELECT_FIELDS = [
  'id',
  'subject',
  'start',
  'end',
  'location',
  'organizer',
  'attendees',
  'isOnlineMeeting',
  'onlineMeeting',
  'isAllDay',
  'isCancelled',
  'webLink',
].join(',');

/**
 * Normalise a user-supplied date/datetime into the form Graph accepts. Port
 * of the Python `_iso8601` helper.
 *
 * Accepts:
 *   - `2026-04-15`           → `2026-04-15T00:00:00`
 *   - `2026-04-15T09:00`     → `2026-04-15T09:00:00`
 *   - `2026-04-15T09:00:00`  → unchanged
 *   - `2026-04-15T09:00:00Z` → `2026-04-15T09:00:00` (trailing Z stripped)
 *
 * Graph carries the timezone separately in `DateTimeTimeZone.timeZone`, so we
 * return an ISO string without any `Z` or offset suffix.
 */
export function normaliseIso(s: string): string {
  let v = s;
  while (v.endsWith('Z') || v.endsWith('z')) {
    v = v.slice(0, -1);
  }
  if (!v.includes('T')) {
    v += 'T00:00:00';
  } else {
    // Count colons after the 'T' to decide if seconds are present.
    const tIdx = v.indexOf('T');
    const timePart = v.slice(tIdx + 1);
    const colons = (timePart.match(/:/g) ?? []).length;
    if (colons === 1) v += ':00';
  }
  return v;
}

function attendeeSummaryOf(a: Attendee): AttendeeSummary {
  const ea = a.emailAddress ?? null;
  const status = a.status ?? null;
  return {
    address: ea?.address ?? null,
    name: ea?.name ?? null,
    type: a.type ?? null,
    response: status?.response ?? null,
  };
}

function pickTime(d: DateTimeTimeZone | null | undefined): {
  dateTime: string | null;
  timeZone: string | null;
} {
  if (!d) return { dateTime: null, timeZone: null };
  return { dateTime: d.dateTime ?? null, timeZone: d.timeZone ?? null };
}

/** Mirror of the Python `_event_summary` helper. */
export function flattenEvent(e: Event): EventSummary {
  const start = pickTime(e.start);
  const end = pickTime(e.end);
  const organizer = e.organizer?.emailAddress?.address ?? null;
  return {
    id: e.id ?? null,
    subject: e.subject ?? null,
    start: start.dateTime,
    end: end.dateTime,
    timeZone: start.timeZone,
    location: locationDisplayName(e.location),
    organizer,
    attendees: (e.attendees ?? []).map(attendeeSummaryOf),
    isOnlineMeeting: e.isOnlineMeeting ?? null,
    onlineJoinUrl: e.onlineMeeting?.joinUrl ?? null,
    isAllDay: e.isAllDay ?? null,
    isCancelled: e.isCancelled ?? null,
    webLink: e.webLink ?? null,
  };
}

function locationDisplayName(
  l: GraphLocation | null | undefined,
): string | null {
  if (!l) return null;
  return l.displayName ?? null;
}

function flattenEventDetail(e: Event): EventDetail {
  const base = flattenEvent(e);
  return {
    ...base,
    bodyContentType: e.body?.contentType ?? null,
    body: e.body?.content ?? null,
    bodyPreview: e.bodyPreview ?? null,
  };
}

function nowUtcIso(): string {
  // Drop millis (Graph's calendarView accepts second-precision ISO + Z).
  return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function plusDaysIso(days: number): string {
  const t = Date.now() + days * 86_400_000;
  return new Date(t).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * Format a Date as `YYYY-MM-DD` in the given IANA TZ. Used by
 * `availability` to anchor the window on the calendar day in the user's
 * TZ — not the UTC day — so "today" means the day the user is actually
 * living.
 *
 * Uses `Intl.DateTimeFormat` (`en-CA` gives ISO-shaped output) so this
 * works without a TZ library.
 */
export function dateStringInTz(date: Date, tz: string): string {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: tz,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date);
  const y = parts.find((p) => p.type === 'year')?.value ?? '';
  const m = parts.find((p) => p.type === 'month')?.value ?? '';
  const d = parts.find((p) => p.type === 'day')?.value ?? '';
  return `${y}-${m}-${d}`;
}

/**
 * Add `days` to a `YYYY-MM-DD` string and return the result in the same
 * shape. Negative `days` walks backwards. Calendar-day exact — handles
 * month/year rollover correctly via UTC anchoring (we don't care about
 * DST here because we're adding whole days to a date-only string).
 */
export function addDaysToDateStr(dateStr: string, days: number): string {
  const d = new Date(`${dateStr}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function envelopeOf<T>(raw: RawPageResponse<T>): PageEnvelope<T> {
  const results = raw.value ?? [];
  return {
    results,
    count: results.length,
    nextLink: raw['@odata.nextLink'] ?? null,
  };
}

export class CalendarResource {
  constructor(private readonly graph: Client) {}

  /**
   * GET /me/calendarView — events in a date range. Uses calendarView so
   * recurring events are expanded into individual occurrences.
   */
  async list(options: ListEventsOptions = {}): Promise<PageEnvelope<EventSummary>> {
    const limit = options.limit ?? 50;
    const startDateTime = options.after ? normaliseIso(options.after) + 'Z' : nowUtcIso();
    const endDateTime = options.before
      ? normaliseIso(options.before) + 'Z'
      : plusDaysIso(7);
    try {
      const raw = (await this.graph
        .api('/me/calendarView')
        .query({
          startDateTime,
          endDateTime,
          $top: limit,
          $orderby: 'start/dateTime',
          $select: EVENT_SELECT_FIELDS,
        })
        .get()) as RawPageResponse<Event>;
      const events = raw.value ?? [];
      return {
        results: events.map(flattenEvent),
        count: events.length,
        nextLink: raw['@odata.nextLink'] ?? null,
      };
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /** GET /me/events/<id> — single event in full (with body). */
  async get(eventId: string, options: GetEventOptions = {}): Promise<EventDetail> {
    try {
      let req = this.graph.api(`/me/events/${encodeURIComponent(eventId)}`);
      if (options.preferText) {
        req = req.header('Prefer', 'outlook.body-content-type="text"');
      }
      const e = (await req.get()) as Event;
      return flattenEventDetail(e);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * POST /me/calendar/getSchedule — free/busy view across one or more users.
   *
   * `availabilityView` is a string of digits with one character per
   * `interval`-minute block. Legend: see {@link AVAILABILITY_LEGEND}.
   */
  async availability(input: AvailabilityInput): Promise<AvailabilityResult> {
    if (!Array.isArray(input.emails) || input.emails.length === 0) {
      throw new Error('calendar availability: at least one --emails value is required.');
    }
    const days = input.days ?? 7;
    const interval = input.interval ?? 30;
    const tz = input.timeZone ?? 'UTC';
    const direction = input.direction ?? 'asc';

    // Anchor the window on a calendar day in the user's TZ. Take the date
    // portion of `pivot` (if given) or today-in-tz (if not). Time is always
    // midnight; we use whole-day windows so the result lines up with real
    // calendar days.
    const pivotDateStr = (input.pivot ?? dateStringInTz(new Date(), tz)).slice(0, 10);

    // Window math: both directions include the pivot day. `asc` walks N
    // days forward from pivot; `desc` walks N-1 days backward, ending
    // after the pivot day.
    const startDateStr =
      direction === 'asc'
        ? pivotDateStr
        : addDaysToDateStr(pivotDateStr, -(days - 1));
    const endDateStr =
      direction === 'asc'
        ? addDaysToDateStr(pivotDateStr, days)
        : addDaysToDateStr(pivotDateStr, 1);

    const startTime = `${startDateStr}T00:00:00`;
    const endTime = `${endDateStr}T00:00:00`;

    const payload = {
      schedules: input.emails,
      startTime: { dateTime: startTime, timeZone: tz },
      endTime: { dateTime: endTime, timeZone: tz },
      availabilityViewInterval: interval,
    };

    try {
      // `Prefer: outlook.timezone` makes Graph return scheduleItem datetimes
      // in our chosen TZ, so the CLI doesn't have to convert. Quoting per
      // RFC 7240.
      const resp = (await this.graph
        .api('/me/calendar/getSchedule')
        .header('Prefer', `outlook.timezone="${tz}"`)
        .post(payload)) as RawPageResponse<ScheduleInformation>;
      const schedules = (resp.value ?? []).map((s) => scheduleSummaryOf(s));
      return {
        emails: input.emails,
        startTime,
        endTime,
        timeZone: tz,
        interval,
        schedules,
      };
    } catch (err) {
      throw liftGraphError(err);
    }
  }
}

function scheduleSummaryOf(s: ScheduleInformation): AvailabilityScheduleSummary {
  return {
    scheduleId: s.scheduleId ?? null,
    availabilityView: s.availabilityView ?? null,
    scheduleItems: (s.scheduleItems ?? []).map(scheduleItemOf),
    workingHours: s.workingHours ?? null,
  };
}

function scheduleItemOf(item: ScheduleItem): {
  subject: string | null;
  start: string | null;
  end: string | null;
  status: string | null;
  location: string | null;
} {
  return {
    subject: item.subject ?? null,
    start: item.start?.dateTime ?? null,
    end: item.end?.dateTime ?? null,
    status: item.status ?? null,
    location: item.location ?? null,
  };
}

export { envelopeOf as _envelopeOf };
