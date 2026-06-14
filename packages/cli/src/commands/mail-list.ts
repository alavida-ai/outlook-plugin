import { parseArgs } from 'node:util';

import type { MessageSummary } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail list [options]

List messages in a mail folder (default: inbox).

Options:
  -n, --limit N        Max messages (default: 10).
  -f, --folder NAME    Folder name (inbox, sentitems, drafts, ...) or id.
  -u, --unread         Only unread messages.
      --from ADDR      Filter by sender address.
      --after DATE     Only messages on/after DATE (YYYY-MM-DD or ISO 8601).
      --before DATE    Only messages on/before DATE (YYYY-MM-DD or ISO 8601).
      --focused        Only Focused Inbox messages.
      --other          Only Other (non-Focused) messages.
      --json           Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        limit: { type: 'string', short: 'n' },
        folder: { type: 'string', short: 'f' },
        unread: { type: 'boolean', short: 'u', default: false },
        from: { type: 'string' },
        after: { type: 'string' },
        before: { type: 'string' },
        focused: { type: 'boolean', default: false },
        other: { type: 'boolean', default: false },
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
    });
  } catch (err) {
    eprintln(formatError(err));
    eprintln(HELP);
    return 1;
  }

  if (parsed.values.help) {
    process.stdout.write(HELP);
    return 0;
  }

  let limit = 10;
  if (parsed.values.limit !== undefined) {
    const n = Number.parseInt(parsed.values.limit, 10);
    if (!Number.isFinite(n) || n <= 0) {
      eprintln(`Invalid --limit: ${parsed.values.limit}`);
      return 1;
    }
    limit = n;
  }

  if (parsed.values.focused && parsed.values.other) {
    eprintln('--focused and --other are mutually exclusive.');
    return 1;
  }

  const ctx = makeContext();
  try {
    const page = await ctx.outlook.mail.list({
      limit,
      folder: parsed.values.folder,
      unread: parsed.values.unread,
      from: parsed.values.from,
      after: parsed.values.after,
      before: parsed.values.before,
      focused: parsed.values.focused,
      other: parsed.values.other,
    });
    // One batch translate covers every row's inbox URL.
    const restIds = page.results
      .map((m) => m.id)
      .filter((id): id is string => typeof id === 'string' && id.length > 0);
    let inboxLinks: Record<string, string | null> = {};
    if (restIds.length > 0) {
      try {
        inboxLinks = await ctx.outlook.mail.inboxLinks(restIds);
      } catch {
        // Translation failures shouldn't block the listing.
      }
    }
    if (parsed.values.json) {
      printJson({
        results: page.results.map((m) => messageSummaryJson(m, inboxLinks[m.id ?? ''] ?? null)),
        count: page.count,
        nextLink: page.nextLink,
      });
    } else {
      renderMessageList(page.results, parsed.values.folder ?? 'inbox', inboxLinks);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function messageSummaryJson(
  m: MessageSummary,
  inboxLink: string | null,
): Record<string, unknown> {
  return {
    id: m.id ?? null,
    subject: m.subject ?? null,
    from: m.from?.emailAddress?.address ?? null,
    receivedDateTime: m.receivedDateTime ?? null,
    isRead: m.isRead ?? null,
    isDraft: m.isDraft ?? null,
    hasAttachments: m.hasAttachments ?? null,
    bodyPreview: m.bodyPreview ?? null,
    // `webLink` is Graph's OWA single-item URL. `inboxLink` opens the new
    // Outlook web app on the inbox with this message selected (built by
    // translating the REST id to `restImmutableEntryId`).
    webLink: m.webLink ?? null,
    inboxLink,
  };
}

function renderMessageList(
  messages: MessageSummary[],
  folder: string,
  inboxLinks: Record<string, string | null>,
): void {
  if (messages.length === 0) {
    println(`(no messages in ${folder})`);
    return;
  }
  println(`${folder} (top ${messages.length})`);
  for (const m of messages) {
    const when = m.receivedDateTime ? m.receivedDateTime.replace('T', ' ').slice(0, 16) : '';
    const from = m.from?.emailAddress?.address ?? '';
    const subject = m.subject ?? '(no subject)';
    const attach = m.hasAttachments ? ' [att]' : '';
    const unread = m.isRead === false ? '* ' : '  ';
    println(`${unread}${when}  ${from}  ${subject}${attach}`);
    if (m.id) println(`    id:   ${m.id}`);
    const link = m.id ? inboxLinks[m.id] : null;
    if (link) println(`    open: ${link}`);
  }
}
