import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail delete <message-id> [options]

Soft-delete a message — Outlook moves it to Deleted Items. Recoverable
until the user empties that folder.

Without --force, prompts for y/N confirmation on stderr. Stdin must
provide the answer; piped invocations should pass --force explicitly.

Options:
      --force          Skip the y/N confirmation prompt.
      --account UPN    Pick a specific cached account.
      --json           Emit JSON envelope instead of human summary.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        force: { type: 'boolean', default: false },
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

  if (!parsed.values.force) {
    const confirmed = await confirm(`Soft-delete message ${messageId}? [y/N] `);
    if (!confirmed) {
      eprintln('Cancelled.');
      return 1;
    }
  }

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const result = await ctx.outlook.mail.delete(messageId);
    if (parsed.values.json) {
      printJson(result);
    } else {
      println(`Deleted ${result.id} (moved to Deleted Items).`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

/**
 * Prompt for y/N confirmation on stderr. Reads a single line from stdin
 * (whether a TTY or piped). Accepts y/Y/yes/Yes; anything else cancels.
 */
async function confirm(prompt: string): Promise<boolean> {
  process.stderr.write(prompt);
  const line = await readLine();
  return /^(y|yes)$/i.test(line.trim());
}

async function readLine(): Promise<string> {
  process.stdin.setEncoding('utf8');
  let buf = '';
  for await (const chunk of process.stdin as AsyncIterable<string>) {
    buf += chunk;
    const nl = buf.indexOf('\n');
    if (nl !== -1) return buf.slice(0, nl);
  }
  return buf;
}
