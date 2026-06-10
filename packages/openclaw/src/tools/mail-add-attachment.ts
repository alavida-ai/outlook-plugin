/**
 * `mail_add_attachment` — attach a local file to an existing draft.
 *
 * Reads the file from the OpenClaw host filesystem (same constraint as
 * `mail_download_attachment`'s `targetPath`), guesses the MIME type from
 * the extension if not provided, and POSTs an inline fileAttachment to
 * Graph. Cap is 3 MB; larger uploads require createUploadSession (future).
 */
import { readFile } from 'node:fs/promises';
import { basename, extname } from 'node:path';
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  draftId: Type.String({
    description: 'Graph draft message id (from `mail_draft` / `mail_reply` / `mail_forward`).',
  }),
  path: Type.String({
    description:
      'Absolute path to the file on the OpenClaw host filesystem. Must be readable. Cap: 3 MB.',
  }),
  name: Type.Optional(
    Type.String({
      description: 'Override the displayed filename. Defaults to the file basename.',
    }),
  ),
  contentType: Type.Optional(
    Type.String({
      description:
        'Override the MIME type. Defaults to a tiny extension-based lookup, then application/octet-stream.',
    }),
  ),
  inline: Type.Optional(
    Type.Boolean({
      description: 'Mark the attachment as inline (referenced by Content-ID in the body).',
    }),
  ),
});

/**
 * Tiny ext → MIME lookup. Same table the CLI uses; duplicated here to avoid
 * making the openclaw package depend on the CLI package. A polish-slice DRY
 * pass would lift it into `@alavida-ai/outlook-core` if it grows further.
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

function guessContentType(filename: string): string {
  const ext = extname(filename).toLowerCase();
  return MIME_BY_EXT[ext] ?? 'application/octet-stream';
}

const mailAddAttachment = defineTool({
  name: 'outlook_mail_add_attachment',
  description:
    "Attach a local file to an existing draft. Reads the file from the OpenClaw host filesystem. Returns the attachment id and metadata. Caps at 3 MB; the upload-session API for larger files is a future enhancement.",
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const bytes = await readFile(params.path);
    const displayName = params.name ?? basename(params.path);
    const contentType = params.contentType ?? guessContentType(displayName);
    const summary = await client.mail.addAttachment(params.draftId, {
      name: displayName,
      contentType,
      contentBytes: bytes,
      isInline: params.inline,
    });
    return {
      ...summary,
      draftId: params.draftId,
    };
  },
});

export default mailAddAttachment;
