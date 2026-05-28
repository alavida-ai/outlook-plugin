import { parseArgs } from 'node:util';

import type { AttachmentSummary } from '@alavida-ai/outlook-core';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail attachments <message-id> [options]

List attachment metadata on a message (id, name, size, content-type, inline?).
Use \`outlook mail download-attachment\` to fetch the bytes.

Options:
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
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
    const page = await ctx.outlook.mail.listAttachments(messageId);
    if (parsed.values.json) {
      printJson({
        results: page.results.map(attachmentJson),
        count: page.count,
        nextLink: page.nextLink,
      });
    } else {
      renderAttachments(messageId, page.results);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function attachmentJson(a: AttachmentSummary): Record<string, unknown> {
  return {
    id: a.id ?? null,
    name: a.name ?? null,
    contentType: a.contentType ?? null,
    size: a.size ?? null,
    isInline: a.isInline ?? false,
  };
}

function renderAttachments(messageId: string, attachments: AttachmentSummary[]): void {
  if (attachments.length === 0) {
    println(`(no attachments on ${messageId})`);
    return;
  }
  println(`Attachments on ${messageId} (${attachments.length})`);
  for (const a of attachments) {
    const size = a.size ?? 0;
    const inline = a.isInline ? ' (inline)' : '';
    println(
      `  ${a.id ?? '(no-id)'}  ${a.name ?? '(no-name)'}  ${size}B  ${a.contentType ?? ''}${inline}`,
    );
  }
}
