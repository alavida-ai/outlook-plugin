# Calendar

Events with attendees send invites the moment they're created — there is no calendar-draft mode equivalent to mail drafts. Confirm attendees + timing with the user before calling `calendar_create` with `attendees: [...]`. See [`./safety.md`](./safety.md) for the confirmation rules.

For passing multi-line `body` (event description) or `comment` (response note) — pass the string with real newlines in your JSON. See [`./body-input.md`](./body-input.md). Calendar event bodies default to HTML; set `bodyContentType: 'Text'` if you want plain text.

## List

```ts
calendar_list({})                                                       // next 7 days (default)
calendar_list({ after: '2026-05-01', before: '2026-05-08' })            // explicit window
calendar_list({ after: '2026-05-01T09:00', before: '2026-05-01T17:00' })// ISO 8601 also OK
calendar_list({ limit: 100 })                                           // raise cap (default 50, max 500)
```

`calendar_list` uses Graph's `calendarView`, which **expands recurring events** to individual occurrences. Each entry in the result is a single occurrence, even for recurring meetings. Window defaults: `after` = now, `before` = now + 7 days.

Return shape:
```json
{ "events": [...], "count": N, "nextLink": null }
```

Per-event fields (subset — see Graph `event` resource for the full surface): `id`, `subject`, `start`, `end`, `location`, `organizer`, `attendees`, `isOnlineMeeting`, `onlineMeeting`, `isAllDay`, `isCancelled`, `webLink`, plus a `body` preview when present.

## Show one event

```ts
calendar_show({ eventId: '<event-id>' })
```

Returns the full event including HTML body, attendee response statuses, and any Teams join URL.

## Create

```ts
calendar_create({
  subject: 'Quick sync',
  start: '2026-05-01T14:00',
  end: '2026-05-01T14:30',
  timeZone: 'Europe/London'
})
```

```ts
// With attendees — sends invites immediately
calendar_create({
  subject: 'Project kickoff',
  start: '2026-05-01T10:00',
  end: '2026-05-01T11:00',
  attendees: ['alice@example.com', 'bob@example.com'],
  location: 'Conference Room A',
  body: 'Agenda: ...'
})
```

```ts
// Teams online meeting
calendar_create({
  subject: 'Remote sync',
  start: '2026-05-01T15:00',
  end: '2026-05-01T15:30',
  isOnlineMeeting: true
})
// Result includes the Teams join URL.
```

```ts
// All-day event
calendar_create({
  subject: 'Offsite',
  start: '2026-06-01',
  end: '2026-06-02',
  isAllDay: true
})
```

```ts
// Recurring (presets: 'daily' | 'weekdays' | 'weekly' | 'monthly' | 'yearly')
calendar_create({
  subject: 'Weekly standup',
  start: '2026-05-04T09:00',
  end: '2026-05-04T09:15',
  recurrence: 'weekdays'
})
```

Time zones: pass an IANA name as `timeZone` (default `UTC`). Same tz applied to `start` and `end`.

Body content type: `body` is treated as HTML by default. For a plain-text description, pass `bodyContentType: 'Text'`.

## Update (PATCH — only the fields you pass)

```ts
calendar_update({ eventId: '<id>', subject: 'New subject' })
calendar_update({ eventId: '<id>', start: '2026-05-01T15:00', end: '2026-05-01T15:30' })
calendar_update({ eventId: '<id>', location: 'Room B' })
calendar_update({ eventId: '<id>', body: 'Updated agenda' })
```

`start` and `end` must be supplied together if either is touched — Graph rejects partial start/end updates. If the event has attendees, updating start/end triggers an invite-update notification to all of them.

## Delete

```ts
calendar_delete({ eventId: '<id>' })
```

Deletion of an event with attendees fires a cancellation notification to all of them. Real-world consequence — confirm with the user before calling. The plugin does not expose an agent-side confirmation prompt; the agent is responsible for the gate.

## Respond to incoming invites

```ts
calendar_respond({ eventId: '<id>', response: 'accept' })
calendar_respond({ eventId: '<id>', response: 'decline', comment: "Conflict — let's reschedule" })
calendar_respond({ eventId: '<id>', response: 'tentative', sendResponse: false })   // don't notify organiser
```

Default behaviour notifies the organiser of your response. `sendResponse: false` suppresses that notification (the organiser's view of your status updates next time they refresh the event).

## Free/busy across users

```ts
calendar_availability({
  emails: ['a@b.com', 'c@d.com'],
  days: 7,
  interval: 60
})
```

```ts
// Higher resolution (30-min blocks), 5 days
calendar_availability({
  emails: ['alice@example.com', 'bob@example.com'],
  days: 5,
  interval: 30
})
```

Defaults: `days: 7`, `interval: 30`, `timeZone: 'UTC'`. Returns a compact "availabilityView" string per user — each character is one block:

| Char | Meaning |
| --- | --- |
| `0` | Free |
| `1` | Tentative |
| `2` | Busy |
| `3` | Out of office |
| `4` | Working elsewhere |

E.g. `"0000022200002220000000"` means free for the first 5 blocks, busy for 3, free for 4, busy for 3, free again. With a 30-min interval and 1-day window, you get 48 chars per user.

The result also includes a per-user `items` array with `subject`, `start`, `end`, `status` for actual conflicts (when the responding user has chosen to share details).

## Common workflows

### "Am I free at 3pm tomorrow?"

```ts
const list = await calendar_list({
  after: '2026-05-02T15:00',
  before: '2026-05-02T16:00'
});
// list.count === 0 → free; > 0 → check list.events
```

### "Find a 30-min slot for me + Alex tomorrow afternoon"

```ts
const avail = await calendar_availability({
  emails: ['me@example.com', 'alex@example.com'],
  days: 1,
  interval: 30
});
// Look for runs of consecutive '0's across both users' availabilityView strings.
```

### "What meetings do I have this week?"

```ts
const list = await calendar_list({ limit: 100 });
for (const ev of list.events) {
  // ev.start, ev.subject, ev.attendees, ...
}
```

### "Cancel my 3pm with Alex" (confirm first!)

```ts
const list = await calendar_list({
  after: '2026-05-02T15:00',
  before: '2026-05-02T16:00'
});
const ev = list.events[0];
// Surface to user: "About to delete event '<ev.subject>' at <ev.start> with N attendees — confirm?"
// After they confirm:
await calendar_delete({ eventId: ev.id });
```
