/**
 * `outlook contacts list` — stub. Mirrors the Python implementation, which
 * prints a TODO to stderr and exits. A real Graph /me/contacts port lives
 * behind its own ticket.
 */
import { parseArgs } from 'node:util';

import { eprintln, formatError, printJson } from '../output.js';

const HELP = `Usage: outlook contacts list [--json]

Stub. Not yet implemented.
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

  if (parsed.values.json) {
    printJson({ stub: true, message: 'contacts_list is a stub; not yet implemented.' });
  } else {
    eprintln('TODO: contacts list (stub).');
  }
  return 0;
}
