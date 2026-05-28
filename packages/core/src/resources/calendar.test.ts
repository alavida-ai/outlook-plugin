import { describe, expect, it, vi } from 'vitest';
import { Client } from '@microsoft/microsoft-graph-client';

import { OutlookClient } from '../client.js';
import { normaliseIso } from './calendar.js';

interface FakeRequest {
  capturedPaths: string[];
  capturedQueries: Array<Record<string, unknown>>;
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
    capturedPosts: [],
    capturedPatches: [],
    capturedMethods: [],
  };

  const chain = {
    query(q: Record<string, unknown>) {
      calls.capturedQueries.push(q);
      return chain;
    },
    header() {
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
    expect(e.subject).toBe('Standup');
    expect(e.body).toBe('<p>agenda</p>');
    expect(e.bodyContentType).toBe('HTML');
    expect(e.attendees[0]?.response).toBe('accepted');
  });
});

describe('CalendarResource.create', () => {
  it('POSTs /me/events with date/timezone envelope and attendees', async () => {
    const { graph, calls } = fakeGraph([
      {
        id: 'evt-new',
        subject: 'Demo',
        start: { dateTime: '2026-04-15T09:00:00.0000000', timeZone: 'UTC' },
        end: { dateTime: '2026-04-15T10:00:00.0000000', timeZone: 'UTC' },
      },
    ]);
    const out = new OutlookClient(graph);
    const created = await out.calendar.create({
      subject: 'Demo',
      start: '2026-04-15T09:00',
      end: '2026-04-15T10:00',
      attendees: ['a@b.com'],
      location: 'Room 1',
      body: 'agenda',
      bodyContentType: 'Text',
      isOnlineMeeting: true,
    });
    expect(calls.capturedPaths).toEqual(['/me/events']);
    expect(calls.capturedMethods).toEqual(['POST']);
    const body = calls.capturedPosts[0] as Record<string, unknown>;
    expect(body.subject).toBe('Demo');
    expect(body.start).toEqual({ dateTime: '2026-04-15T09:00:00', timeZone: 'UTC' });
    expect(body.end).toEqual({ dateTime: '2026-04-15T10:00:00', timeZone: 'UTC' });
    expect(body.attendees).toEqual([
      { emailAddress: { address: 'a@b.com' }, type: 'required' },
    ]);
    expect(body.location).toEqual({ displayName: 'Room 1' });
    expect(body.body).toEqual({ contentType: 'Text', content: 'agenda' });
    expect(body.isOnlineMeeting).toBe(true);
    expect(body.onlineMeetingProvider).toBe('teamsForBusiness');
    expect(created.id).toBe('evt-new');
  });

  it('builds a weekdays recurrence preset', async () => {
    const { graph, calls } = fakeGraph([{ id: 'evt-rec', subject: 'Daily' }]);
    const out = new OutlookClient(graph);
    await out.calendar.create({
      subject: 'Daily',
      start: '2026-04-15T09:00',
      end: '2026-04-15T09:30',
      recurrence: 'weekdays',
    });
    const body = calls.capturedPosts[0] as Record<string, unknown>;
    expect(body.recurrence).toEqual({
      pattern: {
        type: 'weekly',
        interval: 1,
        daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
      },
      range: { type: 'noEnd', startDate: '2026-04-15' },
    });
  });

  it('builds a monthly recurrence preset with dayOfMonth from start', async () => {
    const { graph, calls } = fakeGraph([{ id: 'evt-rec' }]);
    const out = new OutlookClient(graph);
    await out.calendar.create({
      subject: 'Monthly',
      start: '2026-04-15T09:00',
      end: '2026-04-15T09:30',
      recurrence: 'monthly',
    });
    const body = calls.capturedPosts[0] as Record<string, unknown>;
    expect(body.recurrence).toEqual({
      pattern: { type: 'absoluteMonthly', interval: 1, dayOfMonth: 15 },
      range: { type: 'noEnd', startDate: '2026-04-15' },
    });
  });
});

describe('CalendarResource.update', () => {
  it('PATCHes only the supplied fields', async () => {
    const { graph, calls } = fakeGraph([
      { id: 'evt-1', subject: 'Renamed' },
    ]);
    const out = new OutlookClient(graph);
    await out.calendar.update('evt-1', {
      subject: 'Renamed',
      location: 'Room 2',
    });
    expect(calls.capturedPaths).toEqual(['/me/events/evt-1']);
    expect(calls.capturedMethods).toEqual(['PATCH']);
    expect(calls.capturedPatches[0]).toEqual({
      subject: 'Renamed',
      location: { displayName: 'Room 2' },
    });
  });

  it('PATCHes start + end together', async () => {
    const { graph, calls } = fakeGraph([{ id: 'evt-1' }]);
    const out = new OutlookClient(graph);
    await out.calendar.update('evt-1', {
      start: '2026-04-15T10:00',
      end: '2026-04-15T11:00',
      timeZone: 'America/New_York',
    });
    const patch = calls.capturedPatches[0] as Record<string, unknown>;
    expect(patch.start).toEqual({
      dateTime: '2026-04-15T10:00:00',
      timeZone: 'America/New_York',
    });
    expect(patch.end).toEqual({
      dateTime: '2026-04-15T11:00:00',
      timeZone: 'America/New_York',
    });
  });

  it('rejects partial start/end updates', async () => {
    const { graph } = fakeGraph([]);
    const out = new OutlookClient(graph);
    await expect(
      out.calendar.update('evt-1', { start: '2026-04-15T10:00' }),
    ).rejects.toThrow('start and --end');
  });
});

describe('CalendarResource.delete', () => {
  it('DELETEs /me/events/<id>', async () => {
    const { graph, calls } = fakeGraph([undefined]);
    const out = new OutlookClient(graph);
    const r = await out.calendar.delete('evt-1');
    expect(calls.capturedPaths).toEqual(['/me/events/evt-1']);
    expect(calls.capturedMethods).toEqual(['DELETE']);
    expect(r).toEqual({ id: 'evt-1', deleted: true });
  });
});

describe('CalendarResource.respond', () => {
  it('POSTs accept with comment + sendResponse', async () => {
    const { graph, calls } = fakeGraph([undefined]);
    const out = new OutlookClient(graph);
    const r = await out.calendar.respond('evt-1', {
      response: 'accept',
      comment: 'see you there',
    });
    expect(calls.capturedPaths).toEqual(['/me/events/evt-1/accept']);
    expect(calls.capturedPosts[0]).toEqual({
      sendResponse: true,
      comment: 'see you there',
    });
    expect(r).toEqual({
      id: 'evt-1',
      response: 'accept',
      sentResponse: true,
    });
  });

  it('routes tentative to tentativelyAccept', async () => {
    const { graph, calls } = fakeGraph([undefined]);
    const out = new OutlookClient(graph);
    await out.calendar.respond('evt-1', { response: 'tentative', sendResponse: false });
    expect(calls.capturedPaths).toEqual(['/me/events/evt-1/tentativelyAccept']);
    expect(calls.capturedPosts[0]).toEqual({ sendResponse: false });
  });
});

describe('CalendarResource.availability', () => {
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
});

// Silences a lint warning about importing `vi` when no spies are used in this
// file — kept for parity with mail.test.ts and to make adding spies trivial.
void vi;
