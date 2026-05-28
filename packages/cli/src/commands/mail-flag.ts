import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail flag <message-id> [flagged|complete|notFlagged] [options]

Set the follow-up flag on a message.

Status is case-insensitive: 'flagged', 'complete', or 'notflagged'.
Defaults to 'flagged' when omitted.

Options:
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
`;

const FLAG_MAP: Record<string, 'flagged' | 'complete' | 'notFlagged'> = {
  flagged: 'flagged',
  complete: 'complete',
  notflagged: 'notFlagged',
};

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

  const rawStatus = parsed.positionals[1] ?? 'flagged';
  const status = FLAG_MAP[rawStatus.toLowerCase()];
  if (!status) {
    eprintln('status must be: flagged | complete | notFlagged');
    return 1;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const result = await ctx.outlook.mail.flag(messageId, status);
    if (parsed.values.json) {
      printJson(result);
    } else {
      println(`Flag set ${result.flagStatus}: ${result.id}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
