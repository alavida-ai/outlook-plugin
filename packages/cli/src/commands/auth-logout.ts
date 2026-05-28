import { parseArgs } from 'node:util';
import { logout } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook auth logout [--json]

Remove the cached tokens and lock file. Idempotent.
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
    await logout({ cache: ctx.cache });
    if (parsed.values.json) {
      printJson({ status: 'logged_out' });
    } else {
      println('Logged out.');
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
