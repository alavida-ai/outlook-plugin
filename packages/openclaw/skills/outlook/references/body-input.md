# Body & comment input

How to pass text content into tool `body` / `comment` parameters. Applies to:

- `mail_draft({ body, html? })`
- `mail_reply({ body, html? })`
- `mail_forward({ comment? })` — plain text only
- `mail_add_attachment` — no body, but the attachment can carry one
- `calendar_create({ body?, bodyContentType? })`
- `calendar_update({ body?, bodyContentType? })`
- `calendar_respond({ comment? })`

## Multi-line content

Pass the body as a plain JSON string with **real newlines** embedded in it. There is no shell-escape decoding, no stdin pattern, no body-file path — the tool boundary is JSON, so a `\n` in your JSON source encodes to a literal newline character in the resulting string, which is exactly what you want.

```ts
mail_draft({
  to: ['alice@example.com'],
  subject: 'Status',
  body: 'Hi Alice,\n\nQuick update on the deal:\n  - All docs signed\n  - Closing scheduled for Tuesday\n\nBest,\nAgent'
})
```

```ts
calendar_create({
  subject: 'Project review',
  start: '2026-05-15T10:00',
  end: '2026-05-15T11:00',
  attendees: ['alice@example.com'],
  body: 'Agenda:\n  1. Status update\n  2. Risks + blockers\n  3. Next steps\n\nPre-read in shared drive.'
})
```

## HTML bodies

### Mail

`mail_draft` and `mail_reply` accept `html: true` to send the `body` as HTML. Use it when you need rich formatting (lists, bold, links, tables) that plain text can't carry.

```ts
mail_draft({
  to: ['alice@example.com'],
  subject: 'Update',
  body: '<p>Hi Alice,</p><p>Quick update:</p><ul><li>point 1</li><li>point 2</li></ul><p>Best,<br>Agent</p>',
  html: true
})
```

`mail_forward({ comment })` is always plain text — the comment is prepended above Graph's auto-generated quoted body.

### Calendar

`calendar_create` and `calendar_update` accept `bodyContentType: 'HTML' | 'Text'`. Default is `'HTML'`. Pass `'Text'` if you want a plain-text description.

```ts
calendar_create({
  subject: 'Standup',
  start: '2026-05-15T09:00',
  end: '2026-05-15T09:15',
  body: 'Agenda:\n  - Status\n  - Blockers\n  - Asks',
  bodyContentType: 'Text'
})
```

`calendar_respond({ comment })` is always plain text.

## What does NOT support multi-line content

- `subject` (event/message subject) — single line, no newlines
- `location` (calendar) — single line
- `name` (attachment display name), `contentType` (MIME) — single-token strings
- email addresses — obviously one address per array slot
