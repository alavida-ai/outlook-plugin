import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail importance <message-id> [low|normal|high] [options]

Set the importance level on a message. Defaults to 'normal' when omitted.

Options:
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
`;

const LEVELS = new Set(['low', 'normal', 'high'] as const);
type Level = 'low' | 'normal' | 'high';

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

  const rawLevel = (parsed.positionals[1] ?? 'normal').toLowerCase();
  if (!isLevel(rawLevel)) {
    eprintln('level must be: low | normal | high');
    return 1;
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const result = await ctx.outlook.mail.importance(messageId, rawLevel);
    if (parsed.values.json) {
      printJson(result);
    } else {
      println(`Importance set ${result.importance}: ${result.id}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function isLevel(s: string): s is Level {
  return (LEVELS as ReadonlySet<string>).has(s);
}
