import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';
import { resolveBody } from './mail-draft.js';

const HELP = `Usage: outlook mail reply <message-id> --body BODY [options]

Create a draft reply to a message. Never sends.

Options:
      --body B         Reply body. Interprets \\n, \\r, \\t, \\\\. Use '-' for stdin.
      --body-file PATH Read body from a file (no escape decoding).
      --all            Reply to all recipients on the thread.
      --html           Send body as HTML instead of plain text.
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        body: { type: 'string' },
        'body-file': { type: 'string' },
        all: { type: 'boolean', default: false },
        html: { type: 'boolean', default: false },
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

  let bodyText: string;
  try {
    bodyText = await resolveBody(parsed.values.body, parsed.values['body-file']);
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const summary = await ctx.outlook.mail.reply(messageId, {
      body: bodyText,
      html: parsed.values.html,
      replyAll: parsed.values.all,
    });
    if (parsed.values.json) {
      printJson(summary);
    } else {
      println(`Reply draft created.`);
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
