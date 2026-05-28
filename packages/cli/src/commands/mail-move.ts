import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail move <message-id> <folder> [options]

Move a message to another folder. <folder> accepts a well-known name
(inbox, sentitems, drafts, deleteditems, junkemail, archive, outbox,
scheduled, clutter), a custom folder displayName, or a Graph folder id.

Outlook reassigns ids on move; the new id is reported in the output.

Options:
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
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
  const folder = parsed.positionals[1];
  if (!messageId || !folder) {
    eprintln('Usage: outlook mail move <message-id> <folder>');
    eprintln(HELP);
    return 1;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const result = await ctx.outlook.mail.move(messageId, folder);
    if (parsed.values.json) {
      printJson(result);
    } else {
      println(`Moved ${result.oldId} -> ${result.destinationFolder}`);
      println(`  new id: ${result.id}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
