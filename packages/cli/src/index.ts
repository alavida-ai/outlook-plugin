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
  auth login                 Sign in via device-code flow.
  auth logout                Clear cached tokens.
  auth status                Show the signed-in account, if any.
  whoami                     Print the signed-in user's profile.
  mail list                  List messages in a folder.
  mail read                  Read a single message.
  mail search                Search across all folders (KQL).
  mail folders               List mail folders.
  mail attachments           List attachments on a message.
  mail download-attachment   Save an attachment to disk.
  mail draft                 Create a draft message.
  mail reply                 Create a draft reply (also --all).
  mail forward               Create a draft forward.
  mail add-attachment        Attach a file to a draft.
  calendar list              List events in a date range (calendarView).
  calendar show              Show one event in full.
  calendar availability      Free/busy across one or more users.

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

  if (first === 'mail') {
    if (second === 'list') {
      const { run } = await import('./commands/mail-list.js');
      return run(rest);
    }
    if (second === 'read') {
      const { run } = await import('./commands/mail-read.js');
      return run(rest);
    }
    if (second === 'search') {
      const { run } = await import('./commands/mail-search.js');
      return run(rest);
    }
    if (second === 'folders') {
      const { run } = await import('./commands/mail-folders.js');
      return run(rest);
    }
    if (second === 'attachments' || second === 'list-attachments') {
      const { run } = await import('./commands/mail-attachments.js');
      return run(rest);
    }
    if (second === 'download-attachment') {
      const { run } = await import('./commands/mail-download-attachment.js');
      return run(rest);
    }
    if (second === 'draft') {
      const { run } = await import('./commands/mail-draft.js');
      return run(rest);
    }
    if (second === 'reply') {
      const { run } = await import('./commands/mail-reply.js');
      return run(rest);
    }
    if (second === 'forward') {
      const { run } = await import('./commands/mail-forward.js');
      return run(rest);
    }
    if (second === 'add-attachment') {
      const { run } = await import('./commands/mail-add-attachment.js');
      return run(rest);
    }
    eprintln(`Unknown mail subcommand: ${second ?? '(none)'}.`);
    return 1;
  }

  if (first === 'calendar') {
    if (second === 'list') {
      const { run } = await import('./commands/calendar-list.js');
      return run(rest);
    }
    if (second === 'show') {
      const { run } = await import('./commands/calendar-show.js');
      return run(rest);
    }
    if (second === 'availability') {
      const { run } = await import('./commands/calendar-availability.js');
      return run(rest);
    }
    eprintln(`Unknown calendar subcommand: ${second ?? '(none)'}.`);
    return 1;
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
