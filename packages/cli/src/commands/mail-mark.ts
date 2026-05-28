import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail mark <message-id> (--read | --unread) [options]

Mark a message read or unread. Exactly one of --read / --unread is required.

Options:
      --read           Mark as read.
      --unread         Mark as unread.
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        read: { type: 'boolean', default: false },
        unread: { type: 'boolean', default: false },
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

  const { read, unread } = parsed.values;
  if (read === unread) {
    // Both true or both false — ambiguous.
    eprintln('Exactly one of --read or --unread is required.');
    return 1;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const result = await ctx.outlook.mail.mark(messageId, read);
    if (parsed.values.json) {
      printJson(result);
    } else {
      const state = result.isRead ? 'read' : 'unread';
      println(`Marked ${state}: ${result.id}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
