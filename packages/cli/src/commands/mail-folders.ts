import { parseArgs } from 'node:util';

import type { MailFolderSummary } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail folders [--json]

List the signed-in user's mail folders with item counts.
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
    const page = await ctx.outlook.mail.listFolders();
    if (parsed.values.json) {
      printJson({
        results: page.results.map(folderJson),
        count: page.count,
        nextLink: page.nextLink,
      });
    } else {
      renderFolders(page.results);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function folderJson(f: MailFolderSummary): Record<string, unknown> {
  return {
    id: f.id ?? null,
    displayName: f.displayName ?? null,
    unreadItemCount: f.unreadItemCount ?? 0,
    totalItemCount: f.totalItemCount ?? 0,
  };
}

function renderFolders(folders: MailFolderSummary[]): void {
  if (folders.length === 0) {
    println('(no folders)');
    return;
  }
  println(`Mail folders (${folders.length})`);
  for (const f of folders) {
    const unread = f.unreadItemCount ?? 0;
    const total = f.totalItemCount ?? 0;
    println(`  ${f.displayName ?? '(unnamed)'}  (${unread}/${total})  ${f.id ?? ''}`);
  }
}
