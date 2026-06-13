import { parseArgs } from 'node:util';

import type { MessageSummary } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail search <kql-query> [options]

Search across all mail folders using a KQL query
(e.g. \`from:boss@co.com subject:urgent\`).

Options:
  -n, --limit N        Max results (default: 25).
      --json           Emit JSON envelope.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        limit: { type: 'string', short: 'n' },
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

  const query = parsed.positionals[0];
  if (!query) {
    eprintln('Missing required <kql-query>.');
    eprintln(HELP);
    return 1;
  }

  let limit = 25;
  if (parsed.values.limit !== undefined) {
    const n = Number.parseInt(parsed.values.limit, 10);
    if (!Number.isFinite(n) || n <= 0) {
      eprintln(`Invalid --limit: ${parsed.values.limit}`);
      return 1;
    }
    limit = n;
  }

  const ctx = makeContext();
  try {
    const page = await ctx.outlook.mail.search({ query, limit });
    if (parsed.values.json) {
      printJson({
        results: page.results.map(summaryJson),
        count: page.count,
        nextLink: page.nextLink,
      });
    } else {
      renderSearchResults(query, page.results);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

function summaryJson(m: MessageSummary): Record<string, unknown> {
  return {
    id: m.id ?? null,
    subject: m.subject ?? null,
    from: m.from?.emailAddress?.address ?? null,
    receivedDateTime: m.receivedDateTime ?? null,
    isRead: m.isRead ?? null,
    hasAttachments: m.hasAttachments ?? null,
    bodyPreview: m.bodyPreview ?? null,
    webLink: m.webLink ?? null,
  };
}

function renderSearchResults(query: string, messages: MessageSummary[]): void {
  println(`Search: "${query}" (${messages.length} results)`);
  if (messages.length === 0) return;
  for (const m of messages) {
    const when = m.receivedDateTime ? m.receivedDateTime.replace('T', ' ').slice(0, 16) : '';
    const from = m.from?.emailAddress?.address ?? '';
    println(`  ${when}  ${from}  ${m.subject ?? '(no subject)'}`);
    if (m.id) println(`    id: ${m.id}`);
  }
}
