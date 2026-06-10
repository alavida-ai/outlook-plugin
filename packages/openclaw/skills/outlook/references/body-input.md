# Body & comment input

How to pass text content into tool `body` / `comment` parameters. Applies to:

- `outlook_mail_draft({ body, html? })`
- `outlook_mail_reply({ body, html? })`
- `outlook_mail_forward({ comment? })` — plain text only
- `outlook_mail_add_attachment` — no body, but the attachment itself can carry one

(The calendar surface is read-only; no body/comment input.)

## Multi-line content

Pass the body as a plain JSON string with **real newlines** embedded. There is no shell-escape decoding, no stdin pattern, no body-file path — the tool boundary is JSON, so a `\n` in your JSON source encodes to a literal newline character in the resulting string, which is exactly what you want.

```ts
outlook_mail_draft({
  to: ['alice@example.com'],
  subject: 'Status',
  body: 'Hi Alice,\n\nQuick update on the deal:\n  - All docs signed\n  - Closing scheduled for Tuesday\n\nBest,\nAgent'
})
```

## HTML bodies

`outlook_mail_draft` and `outlook_mail_reply` accept `html: true` to send the `body` as HTML. Use it when you need rich formatting (lists, bold, links, tables) that plain text can't carry.

```ts
outlook_mail_draft({
  to: ['alice@example.com'],
  subject: 'Update',
  body: '<p>Hi Alice,</p><p>Quick update:</p><ul><li>point 1</li><li>point 2</li></ul><p>Best,<br>Agent</p>',
  html: true
})
```

`outlook_mail_forward({ comment })` is always plain text — the comment is prepended above Graph's auto-generated quoted body.

## What does NOT support multi-line content

- `subject` (message subject) — single line, no newlines
- `name` (attachment display name), `contentType` (MIME) — single-token strings
- email addresses — obviously one address per array slot
