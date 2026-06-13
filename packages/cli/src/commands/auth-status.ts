import { parseArgs } from 'node:util';
import { status } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook auth status [--json]

Print the cached account, or exit 1 if not logged in.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
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

  const ctx = makeContext();
  try {
    const account = await status({ app: ctx.app, cache: ctx.cache, preferredUpn: undefined });
    if (!account) {
      if (parsed.values.json) {
        printJson({ logged_in: false });
      } else {
        eprintln('Not logged in. Run `outlook auth login`.');
      }
      return 1;
    }
    if (parsed.values.json) {
      printJson({
        logged_in: true,
        username: account.username,
        tenantId: account.tenantId,
        homeAccountId: account.homeAccountId,
      });
    } else {
      println(`Signed in as ${account.username} (tenant ${account.tenantId}).`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
