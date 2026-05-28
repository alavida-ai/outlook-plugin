import { parseArgs } from 'node:util';

import type { MessageFull } from '@alavida-ai/outlook-core';

import { makeContext, resolveUpn } from '../client.js';

interface RecipientLike {
  emailAddress?: { address?: string | null } | null;
}
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail read <message-id> [options]

Read a single message in full.

Options:
      --text           Request plain-text body (default: HTML).
      --account UPN    Pick a specific cached account.
      --json           Emit full JSON.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        text: { type: 'boolean', default: false },
        account: { type: 'string' },
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
      allowPositionals: true,
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

  const messageId = parsed.positionals[0];
  if (!messageId) {
    eprintln('Missing required <message-id>.');
    eprintln(HELP);
    return 1;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const msg = await ctx.outlook.mail.get(messageId, { preferText: parsed.values.text });
    if (parsed.values.json) {
      printJson(messageFullJson(msg));
    } else {
      renderMessage(msg);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function messageFullJson(m: MessageFull): Record<string, unknown> {
  return {
    id: m.id ?? null,
    subject: m.subject ?? null,
    from: m.from?.emailAddress?.address ?? null,
    to: addresses(m.toRecipients),
    cc: addresses(m.ccRecipients),
    bcc: addresses(m.bccRecipients),
    receivedDateTime: m.receivedDateTime ?? null,
    isRead: m.isRead ?? null,
    hasAttachments: m.hasAttachments ?? null,
    importance: m.importance ?? null,
    bodyContentType: m.body?.contentType ?? null,
    body: m.body?.content ?? null,
    webLink: m.webLink ?? null,
  };
}

function addresses(list: RecipientLike[] | null | undefined): string[] {
  if (!list) return [];
  const out: string[] = [];
  for (const r of list) {
    const a = r.emailAddress?.address;
    if (a) out.push(a);
  }
  return out;
}

function renderMessage(m: MessageFull): void {
  println(`Subject: ${m.subject ?? '(no subject)'}`);
  println(`From:    ${m.from?.emailAddress?.address ?? '(unknown)'}`);
  const to = addresses(m.toRecipients).join(', ');
  if (to) println(`To:      ${to}`);
  const cc = addresses(m.ccRecipients).join(', ');
  if (cc) println(`Cc:      ${cc}`);
  if (m.receivedDateTime) println(`Date:    ${m.receivedDateTime}`);
  if (m.importance && m.importance !== 'normal') {
    println(`Importance: ${m.importance}`);
  }
  println('');
  println(m.body?.content ?? '');
  if (m.webLink) {
    println('');
    println(`Open in Outlook: ${m.webLink}`);
  }
}
