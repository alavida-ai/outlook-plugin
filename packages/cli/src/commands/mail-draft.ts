import { readFile } from 'node:fs/promises';
import { parseArgs } from 'node:util';

import { decodeEscapes } from '../escapes.js';
import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail draft --to ADDR --subject SUBJECT (--body BODY | --body-file PATH | --body -) [options]

Create a draft message in the user's Drafts folder. Never sends.

Options:
      --to ADDR        Recipient email address. Repeatable.
      --cc ADDR        CC address. Repeatable.
      --bcc ADDR       BCC address. Repeatable.
      --subject S      Email subject.
      --body B         Body text. Interprets \\n, \\r, \\t, \\\\. Use '-' for stdin.
      --body-file PATH Read body from a file (no escape decoding).
      --html           Send body as HTML instead of plain text.
      --json           Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        to: { type: 'string', multiple: true },
        cc: { type: 'string', multiple: true },
        bcc: { type: 'string', multiple: true },
        subject: { type: 'string' },
        body: { type: 'string' },
        'body-file': { type: 'string' },
        html: { type: 'boolean', default: false },
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

  const to = parsed.values.to ?? [];
  if (to.length === 0) {
    eprintln('At least one --to is required.');
    return 1;
  }
  if (!parsed.values.subject) {
    eprintln('Missing --subject.');
    return 1;
  }

  let bodyText: string;
  try {
    bodyText = await resolveBody(parsed.values.body, parsed.values['body-file']);
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }

  const ctx = makeContext();
  try {
    const summary = await ctx.outlook.mail.draft({
      subject: parsed.values.subject,
      body: bodyText,
      html: parsed.values.html,
      to,
      cc: parsed.values.cc,
      bcc: parsed.values.bcc,
    });
    if (parsed.values.json) {
      printJson(summary);
    } else {
      println(`Draft "${summary.subject ?? ''}" created.`);
      println(`  id: ${summary.id}`);
      if (summary.composeLink) println(`  edit: ${summary.composeLink}`);
      else if (summary.webLink) println(`  open: ${summary.webLink}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

/**
 * Pick the body source. Mirrors the Python CLI's `_resolve_body`:
 *
 *   - `--body` and `--body-file` are mutually exclusive.
 *   - `--body -` reads stdin to EOF.
 *   - `--body STRING` decodes \n / \r / \t / \\ escapes.
 *   - `--body-file PATH` reads the file as-is (no escape decoding).
 *   - Nothing → error.
 *
 * Exported so `mail reply` / `mail forward` can reuse it.
 */
export async function resolveBody(
  body: string | undefined,
  bodyFile: string | undefined,
): Promise<string> {
  if (body !== undefined && bodyFile !== undefined) {
    throw new Error('--body and --body-file are mutually exclusive.');
  }
  if (body === '-') {
    return readStdin();
  }
  if (body !== undefined) {
    return decodeEscapes(body);
  }
  if (bodyFile !== undefined) {
    return readFile(bodyFile, 'utf8');
  }
  throw new Error('Provide --body, --body-file, or pipe via stdin (--body -).');
}

async function readStdin(): Promise<string> {
  process.stdin.setEncoding('utf8');
  const chunks: string[] = [];
  for await (const chunk of process.stdin as AsyncIterable<string>) {
    chunks.push(chunk);
  }
  return chunks.join('');
}
