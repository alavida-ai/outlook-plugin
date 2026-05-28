#!/usr/bin/env node
/**
 * outlook — Microsoft Outlook CLI on top of @alavida-ai/outlook-core.
 *
 * Stdout = data. Stderr = human messages. Exit 0 success, 1 user/auth
 * error, 2 unexpected error.
 *
 * Uses Node's util.parseArgs — no CLI-framework dep.
 */
import { eprintln, formatError } from './output.js';

const TOP_HELP = `Usage: outlook <command> [args...]

Commands:
  auth login         Sign in via device-code flow.
  auth logout        Clear cached tokens.
  auth status        Show the signed-in account, if any.
  whoami             Print the signed-in user's profile.

Global flags:
  --account UPN      Pick a specific cached account (or set OUTLOOK_ACCOUNT).
  --json             Emit JSON to stdout instead of a human summary.

Environment:
  AZURE_CLIENT_ID    Override the embedded Entra app id.
  AZURE_TENANT_ID    Override the default tenant ('common').
  OUTLOOK_ACCOUNT    Default UPN to use when multiple accounts cached.

Run \`outlook <command> --help\` for command-specific options.
`;

async function main(argv: string[]): Promise<number> {
  const [first, second, ...rest] = argv;

  if (!first || first === '--help' || first === '-h' || first === 'help') {
    process.stdout.write(TOP_HELP);
    return first ? 0 : 1;
  }

  if (first === 'auth') {
    if (second === 'login') {
      const { run } = await import('./commands/auth-login.js');
      return run(rest);
    }
    if (second === 'logout') {
      const { run } = await import('./commands/auth-logout.js');
      return run(rest);
    }
    if (second === 'status') {
      const { run } = await import('./commands/auth-status.js');
      return run(rest);
    }
    eprintln(`Unknown auth subcommand: ${second ?? '(none)'}.`);
    return 1;
  }

  if (first === 'whoami') {
    const { run } = await import('./commands/whoami.js');
    return run([second, ...rest].filter((v): v is string => v !== undefined));
  }

  eprintln(`Unknown command: ${first}. Run \`outlook --help\`.`);
  return 1;
}

main(process.argv.slice(2))
  .then((code) => process.exit(code))
  .catch((err: unknown) => {
    eprintln(formatError(err));
    process.exit(2);
  });
