# Mail (CLI)

Draft-only by design. The CLI cannot send mail — there is no `outlook mail send` command, and `Mail.Send` is not in the requested scope set. Every write produces a draft for the user to review and send in Outlook.

## Safety window — read this first

Inbound non-draft mail received within the last **30 minutes** is invisible to every read command:

- `outlook mail list` hides it server-side (Graph `$filter`).
- `outlook mail search` post-filters it from results — the `count` you see is post-filter.
- `outlook mail read` exits with `Mail blocked: …` and a non-zero exit code.
- `outlook mail reply` / `mail forward` / `mail attachments` / `mail download-attachment` all do a pre-flight check and refuse before performing any mutation or fetching attachment bytes.

This protects one-time passwords, 2FA codes, and security alerts from leaking into the user's terminal. **Drafts are exempt** — the user's own composition is always readable.

When you see `Mail blocked: …`, the error also surfaces:
```
  Received:     2026-06-09T15:00:00.000Z
  Available at: 2026-06-09T15:30:00.000Z
```

Surface that `Available at` timestamp to the user, or suggest they handle the message themselves.

## Read

### List

```bash
outlook mail list                                        # last 10 inbox messages
outlook mail list --limit 25                             # last 25
outlook mail list --unread                               # unread only
outlook mail list --from boss@example.com                # filter by sender
outlook mail list --after 2026-04-20                     # received on/after
outlook mail list --after 2026-04-20 --before 2026-04-25 # range
outlook mail list --focused                              # Focused Inbox only
outlook mail list --other                                # Other (non-focused) only
outlook mail list --folder sentitems                     # different folder
outlook mail list --folder drafts                        # user's drafts (no age filter)
```

Folder names: `inbox`, `sentitems`, `drafts`, `deleteditems`, `junkemail`, `archive`, plus any custom folder display name or Graph folder id.

Pretty output per row:
```
* 2026-06-09 14:20  alice@example.com  Contract review [att]
    id:   AAMkADg3NTE0NWVkLT…
    open: https://outlook.cloud.microsoft/mail/inbox/id/AAQkADg3NTE0NWVkLT…
```

JSON output (`--json`) per row:
```json
{
  "id": "AAMk...", "subject": "...", "from": "...",
  "receivedDateTime": "...", "isRead": false, "isDraft": false,
  "hasAttachments": true, "bodyPreview": "...",
  "webLink": "https://outlook.office365.com/owa/?ItemID=...",
  "inboxLink": "https://outlook.cloud.microsoft/mail/inbox/id/..."
}
```

**Use `inboxLink` when telling the user where to open a message** — opens the new Outlook web app's inbox layout with the message selected. `webLink` is Graph's legacy OWA single-item URL.

### Read full message

```bash
outlook mail read '<message-id>'                         # plain-text body (default)
outlook mail read '<message-id>' --html                  # raw HTML if you really need it (rare)
outlook mail read '<message-id>' --json                  # full JSON
```

Plain text is the default — Graph strips HTML/CSS noise server-side. JSON output includes both `webLink` and `inboxLink`.

### Search (KQL)

```bash
outlook mail search 'from:boss@co.com subject:invoice'
outlook mail search 'received:2026-04-01..2026-04-30 hasattachment:true'
outlook mail search '"quarterly review" AND from:boss@co.com' --limit 10
outlook mail search 'sunglobal OR "sun global"' --json
```

KQL property filters: `from:`, `to:`, `cc:`, `subject:`, `body:`, `attachment:`, `hasattachment:`, `received:`, `size:`, `kind:`. Booleans: `AND`, `OR`, `NOT` (uppercase). Wildcards prefix-only (`pref*`). Phrases: `"exact"` (note the doubled quoting in shell: `'"phrase"'`).

Searches **all folders**, not just inbox. Results are **relevance-ranked** (Graph `$search` is mutually exclusive with `$orderby`), not chronological — use `mail list --after/--before` when chronological matters. Fresh non-draft hits are filtered; `count` is post-filter.

## Write (draft only — never sends)

### New draft

```bash
outlook mail draft \
  --to alice@example.com \
  --subject 'Status update' \
  --body 'Hi Alice,\n\nQuick update on the deal:\n  - All docs signed\n  - Closing scheduled for Tuesday\n\nBest,\nAgent'
```

```bash
# Multiple recipients
outlook mail draft \
  --to alice@example.com --to bob@example.com \
  --cc carol@example.com \
  --subject 'Q3 numbers' \
  --body 'See attached.'
```

```bash
# HTML body
outlook mail draft \
  --to alice@example.com \
  --subject 'Update' \
  --body '<p>Hi Alice,</p><p>Quick update:</p><ul><li>point 1</li><li>point 2</li></ul>' \
  --html
```

Returns:
```
Draft created.
  id:           AAMkADg3NTE0...
  Open to edit: https://outlook.cloud.microsoft/mail/compose/AAMkA...
```

The `composeLink` opens the draft directly in compose mode — **always relay this URL to the user**.

### Reply

```bash
outlook mail reply '<message-id>' --body 'Thanks, will look at this Tuesday.'
outlook mail reply '<message-id>' --body '...' --all       # reply-all
outlook mail reply '<message-id>' --body '<p>HTML</p>' --html
```

Same `composeLink` output. Throws `Mail blocked` if the message you're replying to is < 30 min old and isn't a draft — no draft is created.

### Forward

```bash
outlook mail forward '<message-id>' --to next@example.com --comment 'FYI'
outlook mail forward '<message-id>' --to a@b.com --to c@d.com
```

Graph prepends `comment` above the quoted original. Plain text only. Same quarantine behavior.

### Attach a file to an existing draft

```bash
outlook mail add-attachment '<draft-id>' --path ./contract.pdf
outlook mail add-attachment '<draft-id>' --path ./diagram.png --name architecture.png --inline
```

Cap: 3 MB (Graph's inline-attachment limit; chunked upload not yet supported).

## What the CLI deliberately cannot do

The Sun Global pilot scope excludes these. Don't claim otherwise; don't suggest workarounds:

- **Send mail.** No `outlook mail send`, no `Mail.Send` scope.
- **Move / delete / mark read / flag / set importance.** `Mail.ReadWrite` would permit them but no command exists.
- **Manage contacts.** `Contacts.ReadWrite` not in scope.

If the user wants any of these, tell them: "this CLI is read + draft only by design — do it in Outlook directly."

## Folders + attachments

```bash
outlook mail folders                                       # all folders with unread/total counts

outlook mail attachments '<message-id>'                    # list attachments (id, name, size, type, isInline)

outlook mail download-attachment '<msg-id>' '<att-id>'                  # writes to CWD using sanitised name
outlook mail download-attachment '<msg-id>' '<att-id>' --output ~/Downloads/file.pdf
outlook mail download-attachment '<msg-id>' '<att-id>' --output ~/Downloads/   # directory → name appended
```

The attachment commands run the pre-flight quarantine check — they refuse before downloading bytes if the parent message is < 30 min old.

## Common workflows

### "Show me unread mail from this week"

```bash
outlook mail list --unread --after "$(date -v-7d +%Y-%m-%d)" --limit 50
```

### "Draft a reply to the latest email from X"

```bash
msg_id=$(outlook mail list --from x@example.com --limit 1 --json | jq -r '.results[0].id')
outlook mail reply "$msg_id" --body "Thanks for sending. I'll respond by Friday."
# Relay the printed composeLink to the user.
```

### "Read the latest unread email and summarise it"

```bash
msg_id=$(outlook mail list --unread --limit 1 --json | jq -r '.results[0].id')
outlook mail read "$msg_id"
# Plain text body is the default — clean for summarisation.
```

### "Open this message in Outlook"

When surfacing a message to the user, always prefer `inboxLink` over `webLink`:

```bash
outlook mail read '<id>' --json | jq -r '.inboxLink'
# → https://outlook.cloud.microsoft/mail/inbox/id/AAQkADg3...
```

### "Download all the PDFs from the latest email"

```bash
msg_id='<message-id>'
outlook mail attachments "$msg_id" --json | jq -r '.results[] | select(.contentType=="application/pdf") | .id' | \
  while read att_id; do
    outlook mail download-attachment "$msg_id" "$att_id" --output ~/Downloads/
  done
```

## Output stability for scripting

JSON shapes are stable. List commands always carry `{ results: [...], count, nextLink }`. Single-item commands return the typed object directly. Both `webLink` and `inboxLink` are always present in mail outputs (one may be `null` if translation failed; `webLink` falls back gracefully).
