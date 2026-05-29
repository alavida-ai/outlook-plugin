import { parseArgs } from 'node:util';
import { loginDeviceCode } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook auth login [--json]

Run the Microsoft device-code flow and cache the resulting tokens.

  - URL + code are printed to STDERR on the first line. Forward those
    to a human and wait for sign-in to complete.
  - Tokens are cached at ~/.outlook-plugin/tokens.json (0600) — set
    OUTLOOK_TOKEN_CACHE to override.
  - With --json, the result is emitted to stdout as
    { "account": "<upn>" }.
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
    const result = await loginDeviceCode({
      app: ctx.app,
      cache: ctx.cache,
      onDeviceCode: (info) => {
        // First-line-of-stderr UX preserved: the full message includes URL + code.
        eprintln(info.message);
      },
    });
    if (parsed.values.json) {
      printJson({ account: result.account.username });
    } else {
      println(`Signed in as ${result.account.username}.`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
