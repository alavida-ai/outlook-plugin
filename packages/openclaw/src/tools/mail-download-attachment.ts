/**
 * `mail_download_attachment` — fetch an attachment and save to a local path.
 *
 * The agent can't usefully consume binary data in-band, so this tool writes
 * the decoded bytes to disk and returns `{ path, name, contentType, size }`.
 *
 * `targetPath` is REQUIRED. It may be either:
 *   - a directory (the attachment's sanitised name is appended), or
 *   - a full file path (used verbatim).
 *
 * The resolved path is validated against the chosen base directory to
 * prevent an adversarial attachment name from escaping it.
 */
import { mkdir, stat, writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, join, resolve } from 'node:path';
import { Type, type Static } from 'typebox';

import { sanitiseAttachmentName } from '@alavida-ai/outlook-core';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  messageId: Type.String({
    description: 'Graph message id (from mail_list / mail_search).',
  }),
  attachmentId: Type.String({
    description: 'Attachment id (from mail_list_attachments).',
  }),
  targetPath: Type.String({
    description:
      "Target path. May be an existing directory (file gets the attachment's sanitised name appended) or an explicit file path. Must be writable.",
  }),
});

interface DownloadResult {
  path: string;
  name: string;
  contentType: string | null;
  size: number;
}

const mailDownloadAttachment = defineTool({
  name: 'mail_download_attachment',
  description:
    "Download a FileAttachment to a local path. Returns the saved path + metadata. Pass `targetPath` as either an existing directory or an explicit file path.",
  parameters: Params,
  async execute(params: Static<typeof Params>, config): Promise<DownloadResult> {
    const client = getClient(config);
    const result = await client.mail.downloadAttachment(params.messageId, params.attachmentId);
    const target = await resolveTargetPath(params.targetPath, result.name);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, result.contentBytes, { flag: 'w' });
    return {
      path: target,
      name: result.name,
      contentType: result.contentType,
      size: result.size,
    };
  },
});

async function resolveTargetPath(target: string, attachmentName: string): Promise<string> {
  const safeName = sanitiseAttachmentName(attachmentName);
  const absolute = isAbsolute(target) ? target : resolve(process.cwd(), target);
  let isDir = false;
  try {
    const info = await stat(absolute);
    isDir = info.isDirectory();
  } catch {
    isDir = target.endsWith('/');
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

export default mailDownloadAttachment;
