import { mkdir, stat, writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, join, resolve } from 'node:path';
import { parseArgs } from 'node:util';

import { sanitiseAttachmentName } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail download-attachment <message-id> <attachment-id> [options]

Download a FileAttachment to disk. By default writes to the current working
directory using the attachment's (sanitised) name.

Options:
  -o, --output PATH    File path to write to, OR a directory to write into.
      --json           Emit JSON envelope describing the saved file.
`;

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        output: { type: 'string', short: 'o' },
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

  const [messageId, attachmentId] = parsed.positionals;
  if (!messageId || !attachmentId) {
    eprintln('Usage: outlook mail download-attachment <message-id> <attachment-id>');
    return 1;
  }

  const ctx = makeContext();
  try {
    const result = await ctx.outlook.mail.downloadAttachment(messageId, attachmentId);
    const target = await resolveOutputPath(parsed.values.output, result.name);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, result.contentBytes, { flag: 'w' });

    if (parsed.values.json) {
      printJson({
        path: target,
        name: result.name,
        contentType: result.contentType,
        size: result.size,
      });
    } else {
      println(`Wrote ${result.name} (${result.size} bytes) to ${target}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}

/**
 * Resolve the target path for the download.
 *
 *  - No --output      → cwd/<sanitised-name>
 *  - --output is dir  → <dir>/<sanitised-name>
 *  - --output is file → <output>
 *
 * Either way the final path must not escape the chosen base directory after
 * resolution (defence against an attachment name that re-injects path
 * separators despite sanitisation).
 */
async function resolveOutputPath(output: string | undefined, name: string): Promise<string> {
  const safeName = sanitiseAttachmentName(name);
  if (!output) {
    const base = process.cwd();
    return ensureWithin(base, join(base, safeName));
  }
  const absolute = isAbsolute(output) ? output : resolve(process.cwd(), output);
  let isDir = false;
  try {
    const info = await stat(absolute);
    isDir = info.isDirectory();
  } catch {
    // Path doesn't exist yet — treat trailing-slash hints as a directory.
    isDir = output.endsWith('/');
  }
  if (isDir) {
    return ensureWithin(absolute, join(absolute, safeName));
  }
  return absolute;
}

function ensureWithin(base: string, candidate: string): string {
  const baseResolved = resolve(base);
  const candidateResolved = resolve(candidate);
  const baseWithSep = baseResolved.endsWith('/') ? baseResolved : `${baseResolved}/`;
  if (candidateResolved !== baseResolved && !candidateResolved.startsWith(baseWithSep)) {
    throw new Error(`Refusing to write outside ${baseResolved} (got ${candidateResolved})`);
  }
  return candidateResolved;
}
