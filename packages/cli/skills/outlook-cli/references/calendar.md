# Calendar (CLI)

**Read-only.** The CLI's calendar surface is intentionally restricted to reads — no create / update / delete / respond commands exist, and the scope set is `Calendars.Read` + `Calendars.Read.Shared` (no write permission). If the user asks to schedule, reschedule, or cancel anything, surface the constraint plainly and tell them to do it in Outlook directly.

## List events

```bash
outlook calendar list                                                  # next 7 days (default)
outlook calendar list --after 2026-05-01 --before 2026-05-08           # explicit window
outlook calendar list --after '2026-05-01T09:00' --before '2026-05-01T17:00'
outlook calendar list --limit 100                                      # raise cap (default 50)
outlook calendar list --json                                           # full structured output
```

Uses Graph's `calendarView`, which **expands recurring events** to individual occurrences. Each entry is a single occurrence even for recurring meetings. Default window: now → now + 7 days.

Per-event fields (`--json`): `id`, `subject`, `start`, `end`, `timeZone`, `location`, `organizer`, `attendees` (with response status), `isOnlineMeeting`, `onlineJoinUrl`, `isAllDay`, `isCancelled`, `webLink`.

## Show one event

```bash
outlook calendar show '<event-id>'                                     # plain text body (default)
outlook calendar show '<event-id>' --html                              # raw HTML if you really need it
outlook calendar show '<event-id>' --json                              # full structured payload
```

`outlook calendar show` defaults to **plain text** — calendar invite bodies (Teams meetings especially) are mostly HTML chrome that's noise in a terminal. Plain text gives you join URLs, meeting ID, passcode, etc. without the inline `<div style="…">` mess.

## Free/busy across users

```bash
outlook calendar availability --emails a@b.com --emails c@d.com --days 7 --interval 30
outlook calendar availability --emails alice@example.com --days 5 --interval 60 --tz Europe/London
```

Defaults: `--days 7`, `--interval 30`, `--tz UTC`. Returns a compact `availabilityView` digit string per user — one digit per `interval`-minute block walking forward from "now":

| Digit | Meaning |
| --- | --- |
| `0` | Free |
| `1` | Tentative |
| `2` | Busy |
| `3` | Out of office |
| `4` | Working elsewhere |

String length = `days × (1440 / interval)`. Defaults (7 days, 30 min) → 336 digits per user — basically unreadable as one line. The JSON output also carries `scheduleItems` per user (explicit list of meetings with subjects, start, end, status) and `workingHours` per user.

## Common workflows

### "Am I free at 3pm tomorrow?"

```bash
tomorrow=$(date -v+1d +%Y-%m-%d)
count=$(outlook calendar list --after "${tomorrow}T15:00" --before "${tomorrow}T16:00" --json | jq '.count')
[ "$count" -eq 0 ] && echo "Free at 3pm" || echo "Busy at 3pm"
```

### "What's on my calendar tomorrow?"

```bash
tomorrow=$(date -v+1d +%Y-%m-%d)
day_after=$(date -v+2d +%Y-%m-%d)
outlook calendar list --after "$tomorrow" --before "$day_after"
```

### "Find a 30-min slot for me + Alex tomorrow afternoon"

```bash
outlook calendar availability \
  --emails me@example.com --emails alex@example.com \
  --days 1 --interval 30 --json | \
  jq '.results[] | {email: .scheduleId, view: .availabilityView, items: .scheduleItems}'
# Look for runs of '00' (60 min) across both users.
# scheduleItems gives you the explicit conflicts to compare.
```

### "Show me the body of my 10am meeting"

```bash
event_id=$(outlook calendar list --after '2026-05-02T10:00' --before '2026-05-02T10:30' --json | jq -r '.results[0].id')
outlook calendar show "$event_id"
# Plain text body — Teams join URL, meeting ID, passcode, etc.
```

### "Reschedule my 3pm" / "Cancel the 4pm" / "Accept the kickoff invite"

Not supported by this CLI. Tell the user to do it in Outlook directly. The scope set excludes calendar writes by design.

## Output stability for scripting

`outlook calendar list --json` always returns `{ results: [...], count, nextLink }`.
`outlook calendar show --json` returns the typed event object directly.
`outlook calendar availability --json` returns `{ emails, startTime, endTime, timeZone, interval, schedules: [...] }`.
