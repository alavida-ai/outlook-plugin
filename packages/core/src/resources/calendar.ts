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

/** Preset recurrence keywords surfaced by `calendar create`. */
export const RECURRENCE_PRESETS: ReadonlySet<RecurrencePreset> = new Set([
  'daily',
  'weekdays',
  'weekly',
  'monthly',
  'yearly',
]);

export type RecurrencePreset = 'daily' | 'weekdays' | 'weekly' | 'monthly' | 'yearly';

/** Mirror of the Python `ATTENDEE_RESPONSES` constant. */
export const ATTENDEE_RESPONSES: ReadonlySet<AttendeeResponse> = new Set([
  'accept',
  'decline',
  'tentative',
]);

export type AttendeeResponse = 'accept' | 'decline' | 'tentative';

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

export interface ListEventsOptions {
  /** Window start ISO (inclusive). Default: now. */
  after?: string;
  /** Window end ISO (exclusive). Default: now + 7 days. */
  before?: string;
  /** Max events to return. Default: 50. */
  limit?: number;
}

export interface CreateEventInput {
  subject: string;
  /** ISO 8601 start (e.g. `2026-04-15T09:00`). */
  start: string;
  /** ISO 8601 end. */
  end: string;
  /** Optional attendees (required type). */
  attendees?: string[];
  location?: string;
  /** Body content (HTML by default). */
  body?: string;
  /** Body content-type — defaults to 'HTML'. */
  bodyContentType?: 'HTML' | 'Text';
  isAllDay?: boolean;
  /** Add a Teams meeting link. */
  isOnlineMeeting?: boolean;
  /** Recurrence preset. */
  recurrence?: RecurrencePreset;
  /** IANA timezone for start/end. Defaults to 'UTC'. */
  timeZone?: string;
}

export interface UpdateEventInput {
  subject?: string;
  /** Must be paired with `end` if either is supplied. */
  start?: string;
  end?: string;
  location?: string;
  body?: string;
  bodyContentType?: 'HTML' | 'Text';
  /** IANA timezone for any patched start/end. Defaults to 'UTC'. */
  timeZone?: string;
}

export interface RespondInput {
  response: AttendeeResponse;
  comment?: string;
  /** Notify the organiser. Defaults to true. */
  sendResponse?: boolean;
}

export interface AvailabilityInput {
  emails: string[];
  /** Days forward. Default: 7. */
  days?: number;
  /** Interval in minutes for the availability view. Default: 30. */
  interval?: number;
  /** IANA timezone for the query window. Default: 'UTC'. */
  timeZone?: string;
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

interface RecurrencePattern {
  type:
    | 'daily'
    | 'weekly'
    | 'absoluteMonthly'
    | 'absoluteYearly';
  interval: number;
  daysOfWeek?: string[];
  dayOfMonth?: number;
  month?: number;
}

interface RecurrenceRange {
  type: 'noEnd';
  startDate: string;
}

interface PatternedRecurrence {
  pattern: RecurrencePattern;
  range: RecurrenceRange;
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

/**
 * Day-of-week names Graph expects for the `weekly` recurrence pattern. Index
 * matches JavaScript's `Date.prototype.getUTCDay()` returning 0=Sunday.
 */
const DAYS_OF_WEEK: readonly string[] = [
  'sunday',
  'monday',
  'tuesday',
  'wednesday',
  'thursday',
  'friday',
  'saturday',
];

/** Extract the day-of-week (Graph lower-case form) for an ISO datetime. */
function dayOfWeekOf(iso: string): string {
  // ISO without Z is parsed as local time in JS; force UTC so the mapping is stable.
  const parsed = new Date(iso + 'Z');
  if (Number.isNaN(parsed.getTime())) {
    // Fallback: derive numerically. Doesn't happen in practice but keeps tests
    // honest if the consumer passes garbage.
    return 'monday';
  }
  const idx = parsed.getUTCDay();
  return DAYS_OF_WEEK[idx] ?? 'monday';
}

function buildRecurrence(
  preset: RecurrencePreset,
  startIso: string,
): PatternedRecurrence {
  const startDate = startIso.slice(0, 10);
  const day = Number.parseInt(startIso.slice(8, 10), 10);
  const month = Number.parseInt(startIso.slice(5, 7), 10);
  switch (preset) {
    case 'daily':
      return {
        pattern: { type: 'daily', interval: 1 },
        range: { type: 'noEnd', startDate },
      };
    case 'weekdays':
      return {
        pattern: {
          type: 'weekly',
          interval: 1,
          daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
        },
        range: { type: 'noEnd', startDate },
      };
    case 'weekly':
      return {
        pattern: {
          type: 'weekly',
          interval: 1,
          daysOfWeek: [dayOfWeekOf(startIso)],
        },
        range: { type: 'noEnd', startDate },
      };
    case 'monthly':
      return {
        pattern: { type: 'absoluteMonthly', interval: 1, dayOfMonth: day },
        range: { type: 'noEnd', startDate },
      };
    case 'yearly':
      return {
        pattern: {
          type: 'absoluteYearly',
          interval: 1,
          month,
          dayOfMonth: day,
        },
        range: { type: 'noEnd', startDate },
      };
  }
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
  async get(eventId: string): Promise<EventDetail> {
    try {
      const e = (await this.graph
        .api(`/me/events/${encodeURIComponent(eventId)}`)
        .get()) as Event;
      return flattenEventDetail(e);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * POST /me/events — create an event. Sends invites to attendees by default
   * (this is Graph's behaviour on `/me/events`).
   */
  async create(input: CreateEventInput): Promise<EventDetail> {
    const tz = input.timeZone ?? 'UTC';
    const startIso = normaliseIso(input.start);
    const endIso = normaliseIso(input.end);

    const payload: Record<string, unknown> = {
      subject: input.subject,
      start: { dateTime: startIso, timeZone: tz },
      end: { dateTime: endIso, timeZone: tz },
    };
    if (input.body !== undefined) {
      payload.body = {
        contentType: input.bodyContentType ?? 'HTML',
        content: input.body,
      };
    }
    if (input.location) {
      payload.location = { displayName: input.location };
    }
    if (input.attendees && input.attendees.length > 0) {
      payload.attendees = input.attendees.map((address) => ({
        emailAddress: { address },
        type: 'required',
      }));
    }
    if (input.isAllDay) {
      payload.isAllDay = true;
    }
    if (input.isOnlineMeeting) {
      payload.isOnlineMeeting = true;
      payload.onlineMeetingProvider = 'teamsForBusiness';
    }
    if (input.recurrence) {
      if (!RECURRENCE_PRESETS.has(input.recurrence)) {
        throw new Error(
          `Unknown recurrence preset '${input.recurrence}'. Valid: ${Array.from(RECURRENCE_PRESETS).join(', ')}.`,
        );
      }
      payload.recurrence = buildRecurrence(input.recurrence, startIso);
    }

    try {
      const created = (await this.graph.api('/me/events').post(payload)) as Event;
      return flattenEventDetail(created);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * PATCH /me/events/<id> — patch only the fields the caller supplied.
   *
   * `start` and `end` must be supplied together. Graph rejects partial
   * updates of the start/end pair, so we surface a clear error before
   * making the call.
   */
  async update(eventId: string, input: UpdateEventInput): Promise<EventDetail> {
    const hasStart = input.start !== undefined;
    const hasEnd = input.end !== undefined;
    if (hasStart !== hasEnd) {
      throw new Error(
        'calendar update: --start and --end must be supplied together.',
      );
    }

    const tz = input.timeZone ?? 'UTC';
    const payload: Record<string, unknown> = {};
    if (input.subject !== undefined) payload.subject = input.subject;
    if (hasStart && hasEnd) {
      payload.start = { dateTime: normaliseIso(input.start as string), timeZone: tz };
      payload.end = { dateTime: normaliseIso(input.end as string), timeZone: tz };
    }
    if (input.location !== undefined) {
      payload.location = { displayName: input.location };
    }
    if (input.body !== undefined) {
      payload.body = {
        contentType: input.bodyContentType ?? 'HTML',
        content: input.body,
      };
    }

    try {
      const updated = (await this.graph
        .api(`/me/events/${encodeURIComponent(eventId)}`)
        .patch(payload)) as Event;
      return flattenEventDetail(updated);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * DELETE /me/events/<id> — cancel the event. Notifies attendees if any.
   * Different from mail delete: this is a real cancellation, not a soft-move.
   */
  async delete(eventId: string): Promise<{ id: string; deleted: true }> {
    try {
      await this.graph.api(`/me/events/${encodeURIComponent(eventId)}`).delete();
      return { id: eventId, deleted: true };
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * POST /me/events/<id>/{accept|decline|tentativelyAccept} — reply to an
   * incoming meeting invite. Body carries `{ comment, sendResponse }`.
   */
  async respond(
    eventId: string,
    input: RespondInput,
  ): Promise<{ id: string; response: AttendeeResponse; sentResponse: boolean }> {
    if (!ATTENDEE_RESPONSES.has(input.response)) {
      throw new Error(
        `Unknown response '${input.response}'. Valid: ${Array.from(ATTENDEE_RESPONSES).join(', ')}.`,
      );
    }
    const sendResponse = input.sendResponse ?? true;
    const action =
      input.response === 'accept'
        ? 'accept'
        : input.response === 'decline'
          ? 'decline'
          : 'tentativelyAccept';
    const body: Record<string, unknown> = { sendResponse };
    if (input.comment !== undefined) body.comment = input.comment;
    try {
      await this.graph
        .api(`/me/events/${encodeURIComponent(eventId)}/${action}`)
        .post(body);
      return { id: eventId, response: input.response, sentResponse: sendResponse };
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

    // Graph's getSchedule wants the `dateTime` stripped of any tz suffix —
    // the timezone moves into the `timeZone` field next to it.
    const startTime = nowUtcIso().replace(/Z$/, '');
    const end = new Date(Date.now() + days * 86_400_000).toISOString();
    const endTime = end.replace(/\.\d{3}Z$/, '').replace(/Z$/, '');

    const payload = {
      schedules: input.emails,
      startTime: { dateTime: startTime, timeZone: tz },
      endTime: { dateTime: endTime, timeZone: tz },
      availabilityViewInterval: interval,
    };

    try {
      const resp = (await this.graph
        .api('/me/calendar/getSchedule')
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
