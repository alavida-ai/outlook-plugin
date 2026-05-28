import { parseArgs } from 'node:util';

import type { MessageSummary } from '@alavida-ai/outlook-core';

import { makeContext, resolveUpn } from '../client.js';
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
      --account UPN    Pick a specific cached account.
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
        account: { type: 'string' },
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

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
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
    if (parsed.values.json) {
      printJson({
        results: page.results.map(messageSummaryJson),
        count: page.count,
        nextLink: page.nextLink,
      });
    } else {
      renderMessageList(page.results, parsed.values.folder ?? 'inbox');
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function messageSummaryJson(m: MessageSummary): Record<string, unknown> {
  return {
    id: m.id ?? null,
    subject: m.subject ?? null,
    from: m.from?.emailAddress?.address ?? null,
    receivedDateTime: m.receivedDateTime ?? null,
    isRead: m.isRead ?? null,
    hasAttachments: m.hasAttachments ?? null,
    bodyPreview: m.bodyPreview ?? null,
    webLink: m.webLink ?? null,
  };
}

function renderMessageList(messages: MessageSummary[], folder: string): void {
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
    if (m.id) println(`    id: ${m.id}`);
  }
}
