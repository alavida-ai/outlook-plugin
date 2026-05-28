import { readFile } from 'node:fs/promises';
import { basename, extname } from 'node:path';
import { parseArgs } from 'node:util';

import { makeContext, resolveUpn } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook mail add-attachment <draft-id> --file PATH [options]

Attach a local file to an existing draft.

Options:
      --file PATH       Path to the file to attach (required).
      --name NAME       Override the displayed filename (default: file basename).
      --content-type T  Override the MIME type (default: guessed from extension).
      --inline          Mark the attachment as inline (default: false).
      --account UPN     Pick a specific cached account.
      --json            Emit JSON envelope instead of human summary.

Files are uploaded inline; the cap is 3 MB. Larger uploads require Graph's
createUploadSession flow, which is not yet supported.
`;

/**
 * Tiny extension -> MIME-type map. Falls back to `application/octet-stream`.
 *
 * Deliberately ad-hoc to avoid pulling in `mime-types` as a runtime dep for a
 * dozen extensions. Extend as needed.
 */
const MIME_BY_EXT: Record<string, string> = {
  '.pdf': 'application/pdf',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.txt': 'text/plain',
  '.html': 'text/html',
  '.htm': 'text/html',
  '.csv': 'text/csv',
  '.json': 'application/json',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.doc': 'application/msword',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.zip': 'application/zip',
};

export function guessContentType(filename: string): string {
  const ext = extname(filename).toLowerCase();
  return MIME_BY_EXT[ext] ?? 'application/octet-stream';
}

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        file: { type: 'string' },
        name: { type: 'string' },
        'content-type': { type: 'string' },
        inline: { type: 'boolean', default: false },
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

  const draftId = parsed.positionals[0];
  if (!draftId) {
    eprintln('Missing required <draft-id>.');
    eprintln(HELP);
    return 1;
  }
  if (!parsed.values.file) {
    eprintln('Missing required --file.');
    return 1;
  }

  let bytes: Buffer;
  try {
    bytes = await readFile(parsed.values.file);
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }

  const displayName = parsed.values.name ?? basename(parsed.values.file);
  const contentType = parsed.values['content-type'] ?? guessContentType(displayName);

  const preferredUpn = resolveUpn(parsed.values.account);
  const ctx = makeContext({ preferredUpn });
  try {
    const summary = await ctx.outlook.mail.addAttachment(draftId, {
      name: displayName,
      contentType,
      contentBytes: bytes,
      isInline: parsed.values.inline,
    });
    if (parsed.values.json) {
      printJson(summary);
    } else {
      println(`Attached ${summary.name ?? displayName} (${summary.size ?? bytes.byteLength} bytes) to draft ${draftId}.`);
      if (summary.attachmentId) println(`  attachmentId: ${summary.attachmentId}`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
