import { parseArgs } from 'node:util';

import { decodeEscapes } from '../escapes.js';
import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail forward <message-id> --to ADDR [options]

Create a draft forward of a message. Never sends.

Options:
      --to ADDR        Recipient email address. Repeatable.
      --cc ADDR        CC address. Repeatable.
      --comment TEXT   Optional note prepended above the quoted original.
                       Interprets \\n, \\r, \\t, \\\\ escapes.
      --account UPN    Pick a specific cached account.
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
        comment: { type: 'string' },
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
  const to = parsed.values.to ?? [];
  if (to.length === 0) {
    eprintln('At least one --to is required.');
    return 1;
  }

  const comment =
    parsed.values.comment !== undefined ? decodeEscapes(parsed.values.comment) : undefined;

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const summary = await ctx.outlook.mail.forward(messageId, {
      to,
      cc: parsed.values.cc,
      comment,
    });
    if (parsed.values.json) {
      printJson(summary);
    } else {
      println(`Forward draft created to ${to.join(', ')}.`);
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
