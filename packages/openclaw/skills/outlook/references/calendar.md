# Calendar

**Read-only.** The plugin's calendar surface is intentionally restricted to reads — no create / update / delete / respond tools exist, and the scope set is `Calendars.Read` + `Calendars.Read.Shared` (no write permission). If the user asks the agent to schedule, reschedule, or cancel anything, surface the constraint plainly and ask them to do it in Outlook directly. Don't claim otherwise; don't attempt workarounds.

## List

```ts
outlook_calendar_list({})                                                       // next 7 days (default)
outlook_calendar_list({ after: '2026-05-01', before: '2026-05-08' })            // explicit window
outlook_calendar_list({ after: '2026-05-01T09:00', before: '2026-05-01T17:00' })// ISO 8601 also OK
outlook_calendar_list({ limit: 100 })                                           // raise cap (default 50, max 500)
```

`outlook_calendar_list` uses Graph's `calendarView`, which **expands recurring events** to individual occurrences. Each entry is a single occurrence, even for recurring meetings. Window defaults: `after` = now, `before` = now + 7 days.

Return shape:
```json
{ "events": [...], "count": N, "nextLink": null }
```

Per-event fields: `id`, `subject`, `start`, `end`, `location`, `organizer`, `attendees`, `isOnlineMeeting`, `onlineJoinUrl`, `isAllDay`, `isCancelled`, `webLink`.

## Show one event

```ts
outlook_calendar_show({ eventId: '<event-id>' })                                // plain text body (default)
outlook_calendar_show({ eventId: '<event-id>', preferText: false })             // raw HTML
```

`preferText` defaults to **true** — calendar invite bodies (Teams meetings especially) are mostly HTML chrome that wastes context. Plain text gives you join URLs, meeting ID, passcode, etc. without the inline `<div style="…">` noise. Only pass `preferText: false` if you specifically need the raw HTML.

Returns the full event including body, attendee response statuses, and any Teams join URL.

## Free/busy across users

```ts
outlook_calendar_availability({
  emails: ['a@b.com', 'c@d.com'],
  days: 7,
  interval: 60
})
```

```ts
// Higher resolution (30-min blocks), 5 days
outlook_calendar_availability({
  emails: ['alice@example.com', 'bob@example.com'],
  days: 5,
  interval: 30
})
```

Defaults: `days: 7`, `interval: 30`, `timeZone: 'UTC'`. Returns a compact "availabilityView" string per user — each character is one `interval`-minute block walking forward from `startTime`:

| Char | Meaning |
| --- | --- |
| `0` | Free |
| `1` | Tentative |
| `2` | Busy |
| `3` | Out of office |
| `4` | Working elsewhere |

String length = `days × (1440 / interval)`. So defaults (7 days, 30 min) → 336 chars per user. With 1 day, 30 min → 48 chars (48 blocks of 30 min).

Example `"0000022200002220000000"` (30-min blocks): free for 5 blocks (2.5 h), busy for 3 (1.5 h), free for 4 (2 h), busy for 3 (1.5 h), free again.

The result also includes a per-user `scheduleItems` array with `subject`, `start`, `end`, `status`, `location` for the actual meetings in the window (when that user has chosen to share details). The `availabilityView` digit string is the compressed summary; `scheduleItems` is the explicit list.

`workingHours` per user lets you ignore "free" digits outside their normal hours.

## Common workflows

### "Am I free at 3pm tomorrow?"

```ts
const list = await outlook_calendar_list({
  after: '2026-05-02T15:00',
  before: '2026-05-02T16:00'
});
// list.count === 0 → free; > 0 → check list.events
```

### "What's on my calendar tomorrow?"

```ts
const list = await outlook_calendar_list({
  after: '2026-05-02',
  before: '2026-05-03'
});
// Render list.events as a chronological day summary for the user.
```

### "Find a 30-min slot for me + Alex tomorrow afternoon"

```ts
const avail = await outlook_calendar_availability({
  emails: ['me@example.com', 'alex@example.com'],
  days: 1,
  interval: 30
});
// Look for runs of consecutive '0's across both users' availabilityView strings,
// or compare scheduleItems for explicit conflicts.
```

### "Show me the body of my 10am meeting"

```ts
const list = await outlook_calendar_list({
  after: '2026-05-02T10:00',
  before: '2026-05-02T10:30'
});
const detail = await outlook_calendar_show({ eventId: list.events[0].id });
// detail.body is plain text (Teams join URL, meeting ID, passcode, etc.)
```

### "Reschedule my 3pm" / "Cancel the 4pm" / "Accept the kickoff invite"

Not supported by this plugin. Tell the user to do it in Outlook directly. The scope set excludes calendar writes by design.
