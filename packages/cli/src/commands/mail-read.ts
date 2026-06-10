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
    // Translate this message's REST id to a restImmutableEntryId so we can
    // build the new-Outlook inbox URL alongside Graph's OWA `webLink`.
    let inboxLink: string | null = null;
    if (msg.id) {
      try {
        const links = await ctx.outlook.mail.inboxLinks([msg.id]);
        inboxLink = links[msg.id] ?? null;
      } catch {
        // Translation failures shouldn't block the read — leave inboxLink null.
      }
    }
    if (parsed.values.json) {
      printJson(messageFullJson(msg, inboxLink));
    } else {
      renderMessage(msg, inboxLink);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function messageFullJson(
  m: MessageFull,
  inboxLink: string | null,
): Record<string, unknown> {
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
    // `webLink` is Graph's OWA single-item URL. `inboxLink` opens the new
    // Outlook web app on the inbox with this message selected — built by
    // translating the REST id to `restImmutableEntryId`.
    webLink: m.webLink ?? null,
    inboxLink,
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

function renderMessage(m: MessageFull, inboxLink: string | null): void {
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
  if (inboxLink) {
    println('');
    println(`Open in Outlook: ${inboxLink}`);
  }
}
