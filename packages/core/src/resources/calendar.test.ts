import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Client } from '@microsoft/microsoft-graph-client';

import { OutlookClient } from '../client.js';
import { addDaysToDateStr, dateStringInTz, normaliseIso } from './calendar.js';

interface FakeRequest {
  capturedPaths: string[];
  capturedQueries: Array<Record<string, unknown>>;
  capturedHeaders: Array<[string, string]>;
  capturedPosts: unknown[];
  capturedPatches: unknown[];
  capturedMethods: string[];
}

/**
 * Minimal fake Graph client — chainable api()/query()/get()/post()/patch()/
 * delete(). Mirrors mail.test.ts so the calendar tests stay consistent.
 */
function fakeGraph(responses: unknown[]): { graph: Client; calls: FakeRequest } {
  const queue = [...responses];
  const calls: FakeRequest = {
    capturedPaths: [],
    capturedQueries: [],
    capturedHeaders: [],
    capturedPosts: [],
    capturedPatches: [],
    capturedMethods: [],
  };

  const chain = {
    query(q: Record<string, unknown>) {
      calls.capturedQueries.push(q);
      return chain;
    },
    header(k: string, v: string) {
      calls.capturedHeaders.push([k, v]);
      return chain;
    },
    async get() {
      calls.capturedMethods.push('GET');
      return queue.shift();
    },
    async post(body: unknown) {
      calls.capturedMethods.push('POST');
      calls.capturedPosts.push(body);
      return queue.shift();
    },
    async patch(body: unknown) {
      calls.capturedMethods.push('PATCH');
      calls.capturedPatches.push(body);
      return queue.shift();
    },
    async delete() {
      calls.capturedMethods.push('DELETE');
      return queue.shift();
    },
  };

  const graph = {
    api(path: string) {
      calls.capturedPaths.push(path);
      return chain;
    },
  } as unknown as Client;

  return { graph, calls };
}

describe('normaliseIso', () => {
  it('expands a bare YYYY-MM-DD to midnight', () => {
    expect(normaliseIso('2026-04-15')).toBe('2026-04-15T00:00:00');
  });
  it('appends seconds when missing', () => {
    expect(normaliseIso('2026-04-15T09:00')).toBe('2026-04-15T09:00:00');
  });
  it('strips trailing Z', () => {
    expect(normaliseIso('2026-04-15T09:00:00Z')).toBe('2026-04-15T09:00:00');
    expect(normaliseIso('2026-04-15T09:00Z')).toBe('2026-04-15T09:00:00');
  });
  it('leaves a fully-formed ISO datetime alone', () => {
    expect(normaliseIso('2026-04-15T09:00:00')).toBe('2026-04-15T09:00:00');
  });
});

describe('CalendarResource.list', () => {
  it('hits /me/calendarView with default window', async () => {
    const { graph, calls } = fakeGraph([
      {
        value: [
          {
            id: 'evt-1',
            subject: 'Sync',
            start: { dateTime: '2026-04-15T09:00:00.0000000', timeZone: 'UTC' },
            end: { dateTime: '2026-04-15T09:30:00.0000000', timeZone: 'UTC' },
            organizer: { emailAddress: { address: 'a@b.com', name: 'A' } },
            attendees: [],
            isAllDay: false,
            isCancelled: false,
            isOnlineMeeting: false,
            webLink: 'https://outlook.example/event/evt-1',
          },
        ],
      },
    ]);
    const out = new OutlookClient(graph);
    const page = await out.calendar.list();
    expect(calls.capturedPaths).toEqual(['/me/calendarView']);
    const q = calls.capturedQueries[0] as Record<string, unknown>;
    expect(q.$top).toBe(50);
    expect(q.$orderby).toBe('start/dateTime');
    expect(typeof q.startDateTime).toBe('string');
    expect(typeof q.endDateTime).toBe('string');
    expect(page.count).toBe(1);
    expect(page.results[0]?.subject).toBe('Sync');
    expect(page.results[0]?.organizer).toBe('a@b.com');
  });

  it('passes user-supplied after/before through normaliseIso', async () => {
    const { graph, calls } = fakeGraph([{ value: [] }]);
    const out = new OutlookClient(graph);
    await out.calendar.list({
      after: '2026-04-15',
      before: '2026-04-20T18:00',
      limit: 5,
    });
    const q = calls.capturedQueries[0] as Record<string, unknown>;
    expect(q.startDateTime).toBe('2026-04-15T00:00:00Z');
    expect(q.endDateTime).toBe('2026-04-20T18:00:00Z');
    expect(q.$top).toBe(5);
  });
});

describe('CalendarResource.get', () => {
  it('fetches /me/events/<id> and flattens the body', async () => {
    const { graph, calls } = fakeGraph([
      {
        id: 'evt-1',
        subject: 'Standup',
        start: { dateTime: '2026-04-15T09:00:00.0000000', timeZone: 'UTC' },
        end: { dateTime: '2026-04-15T09:30:00.0000000', timeZone: 'UTC' },
        body: { contentType: 'HTML', content: '<p>agenda</p>' },
        bodyPreview: 'agenda',
        attendees: [
          {
            emailAddress: { address: 'a@b.com', name: 'A' },
            type: 'required',
            status: { response: 'accepted', time: '...' },
          },
        ],
      },
    ]);
    const out = new OutlookClient(graph);
    const e = await out.calendar.get('evt-1');
    expect(calls.capturedPaths).toEqual(['/me/events/evt-1']);
    expect(calls.capturedHeaders).toEqual([]);
    expect(e.subject).toBe('Standup');
    expect(e.body).toBe('<p>agenda</p>');
    expect(e.bodyContentType).toBe('HTML');
    expect(e.attendees[0]?.response).toBe('accepted');
  });

  it('sends Prefer: text body-content-type when preferText is set', async () => {
    const { graph, calls } = fakeGraph([
      {
        id: 'evt-1',
        body: { contentType: 'Text', content: 'agenda (plain text)' },
      },
    ]);
    const out = new OutlookClient(graph);
    const e = await out.calendar.get('evt-1', { preferText: true });
    expect(calls.capturedHeaders).toEqual([
      ['Prefer', 'outlook.body-content-type="text"'],
    ]);
    expect(e.body).toBe('agenda (plain text)');
    expect(e.bodyContentType).toBe('Text');
  });
});

describe('date helpers', () => {
  it('dateStringInTz returns YYYY-MM-DD in the named timezone', () => {
    // 2026-06-11 23:30 UTC is already tomorrow in Sydney.
    const t = new Date('2026-06-11T23:30:00Z');
    expect(dateStringInTz(t, 'UTC')).toBe('2026-06-11');
    expect(dateStringInTz(t, 'Australia/Sydney')).toBe('2026-06-12');
    // And still yesterday in Los Angeles.
    expect(dateStringInTz(t, 'America/Los_Angeles')).toBe('2026-06-11');
  });
  it('addDaysToDateStr walks the calendar correctly', () => {
    expect(addDaysToDateStr('2026-06-11', 0)).toBe('2026-06-11');
    expect(addDaysToDateStr('2026-06-11', 7)).toBe('2026-06-18');
    expect(addDaysToDateStr('2026-06-11', -7)).toBe('2026-06-04');
    // Month rollover
    expect(addDaysToDateStr('2026-06-30', 1)).toBe('2026-07-01');
    expect(addDaysToDateStr('2026-07-01', -1)).toBe('2026-06-30');
    // Year rollover
    expect(addDaysToDateStr('2026-12-31', 1)).toBe('2027-01-01');
  });
});

describe('CalendarResource.availability', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-06-11T12:00:00Z'));
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it('POSTs /me/calendar/getSchedule with the supplied emails', async () => {
    const { graph, calls } = fakeGraph([
      {
        value: [
          {
            scheduleId: 'a@b.com',
            availabilityView: '000222000',
            scheduleItems: [
              {
                subject: 'Lunch',
                start: { dateTime: '2026-04-15T12:00:00' },
                end: { dateTime: '2026-04-15T13:00:00' },
                status: 'busy',
                location: 'Cafe',
              },
            ],
            workingHours: null,
          },
        ],
      },
    ]);
    const out = new OutlookClient(graph);
    const r = await out.calendar.availability({
      emails: ['a@b.com'],
      days: 3,
      interval: 60,
      timeZone: 'America/New_York',
    });
    expect(calls.capturedPaths).toEqual(['/me/calendar/getSchedule']);
    const body = calls.capturedPosts[0] as Record<string, unknown>;
    expect(body.schedules).toEqual(['a@b.com']);
    expect(body.availabilityViewInterval).toBe(60);
    expect((body.startTime as { timeZone: string }).timeZone).toBe(
      'America/New_York',
    );
    expect(r.emails).toEqual(['a@b.com']);
    expect(r.schedules[0]?.availabilityView).toBe('000222000');
    expect(r.schedules[0]?.scheduleItems[0]?.subject).toBe('Lunch');
  });

  it('rejects empty email list', async () => {
    const { graph } = fakeGraph([]);
    const out = new OutlookClient(graph);
    await expect(out.calendar.availability({ emails: [] })).rejects.toThrow(
      'at least one',
    );
  });

  it('asc default: window starts at today (in tz) and runs N days forward', async () => {
    const { graph, calls } = fakeGraph([{ value: [] }]);
    const out = new OutlookClient(graph);
    await out.calendar.availability({
      emails: ['a@b.com'],
      days: 3,
      timeZone: 'Europe/London',
    });
    const body = calls.capturedPosts[0] as Record<string, unknown>;
    // System time pinned at 2026-06-11T12:00:00Z → 13:00 BST → today is 2026-06-11 in London.
    expect((body.startTime as { dateTime: string }).dateTime).toBe(
      '2026-06-11T00:00:00',
    );
    expect((body.endTime as { dateTime: string }).dateTime).toBe(
      '2026-06-14T00:00:00',
    );
    expect((body.startTime as { timeZone: string }).timeZone).toBe('Europe/London');
    // Prefer header anchors response datetimes in the same tz.
    expect(calls.capturedHeaders).toEqual([
      ['Prefer', 'outlook.timezone="Europe/London"'],
    ]);
  });

  it('desc: window ends after pivot and walks back N-1 days', async () => {
    const { graph, calls } = fakeGraph([{ value: [] }]);
    const out = new OutlookClient(graph);
    await out.calendar.availability({
      emails: ['a@b.com'],
      days: 3,
      timeZone: 'UTC',
      direction: 'desc',
    });
    const body = calls.capturedPosts[0] as Record<string, unknown>;
    // today=2026-06-11; desc days=3 → [pivot-2, pivot+1) → 06-09..06-12.
    expect((body.startTime as { dateTime: string }).dateTime).toBe(
      '2026-06-09T00:00:00',
    );
    expect((body.endTime as { dateTime: string }).dateTime).toBe(
      '2026-06-12T00:00:00',
    );
  });

  it('explicit pivot anchors the window without considering "now"', async () => {
    const { graph, calls } = fakeGraph([{ value: [] }]);
    const out = new OutlookClient(graph);
    await out.calendar.availability({
      emails: ['a@b.com'],
      days: 5,
      timeZone: 'UTC',
      pivot: '2026-09-01',
    });
    const body = calls.capturedPosts[0] as Record<string, unknown>;
    expect((body.startTime as { dateTime: string }).dateTime).toBe(
      '2026-09-01T00:00:00',
    );
    expect((body.endTime as { dateTime: string }).dateTime).toBe(
      '2026-09-06T00:00:00',
    );
  });
});
