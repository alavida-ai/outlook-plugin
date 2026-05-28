import type { Client } from '@microsoft/microsoft-graph-client';
import type {
  Attachment,
  FileAttachment,
  MailFolder,
  Message,
} from '@microsoft/microsoft-graph-types';

import { liftGraphError, NotFoundError } from '../graph/errors.js';

/** Well-known Graph folder aliases (case-insensitive on input). */
export const WELL_KNOWN_FOLDERS: ReadonlySet<string> = new Set([
  'inbox',
  'sentitems',
  'drafts',
  'deleteditems',
  'junkemail',
  'archive',
  'outbox',
  'scheduled',
  'clutter',
]);

/** Subset of `Message` we surface in list endpoints. */
export type MessageSummary = Pick<
  Message,
  | 'id'
  | 'subject'
  | 'from'
  | 'receivedDateTime'
  | 'isRead'
  | 'hasAttachments'
  | 'bodyPreview'
  | 'webLink'
>;

/** Full message shape used by `read`. */
export type MessageFull = MessageSummary &
  Pick<
    Message,
    'toRecipients' | 'ccRecipients' | 'bccRecipients' | 'body' | 'importance'
  >;

/** Folder shape returned by `listFolders`. */
export type MailFolderSummary = Pick<
  MailFolder,
  'id' | 'displayName' | 'unreadItemCount' | 'totalItemCount'
>;

/** Attachment metadata shape returned by `listAttachments`. */
export type AttachmentSummary = Pick<
  Attachment,
  'id' | 'name' | 'contentType' | 'size' | 'isInline'
>;

/** Common envelope for paged list results. */
export interface PageEnvelope<T> {
  results: T[];
  count: number;
  nextLink: string | null;
}

export interface ListMessagesOptions {
  limit?: number;
  folder?: string;
  unread?: boolean;
  from?: string;
  after?: string;
  before?: string;
  focused?: boolean;
  other?: boolean;
}

export interface SearchMessagesOptions {
  query: string;
  limit?: number;
}

export interface ReadMessageOptions {
  preferText?: boolean;
}

export interface DownloadAttachmentResult {
  /** Decoded raw bytes. */
  contentBytes: Buffer;
  /** Sanitised attachment name. */
  name: string;
  /** MIME type, if Graph reported one. */
  contentType: string | null;
  /** Byte length after decode. */
  size: number;
}

const MESSAGE_SELECT_FIELDS = [
  'id',
  'subject',
  'from',
  'receivedDateTime',
  'isRead',
  'hasAttachments',
  'bodyPreview',
  'webLink',
].join(',');

const MESSAGE_FULL_SELECT_FIELDS = [
  MESSAGE_SELECT_FIELDS,
  'toRecipients',
  'ccRecipients',
  'bccRecipients',
  'body',
  'importance',
].join(',');

const FOLDER_SELECT_FIELDS = 'id,displayName,unreadItemCount,totalItemCount';

const ATTACHMENT_SELECT_FIELDS = 'id,name,contentType,size,isInline';

/** Heuristic: long opaque strings are Graph folder ids. */
function looksLikeFolderId(s: string): boolean {
  return s.length > 20 && !s.includes(' ');
}

/** Graph wants ISO-8601 with Z suffix in `$filter`. Expand bare dates. */
export function normaliseDateForFilter(raw: string): string {
  let s = raw.trim();
  if (!s.includes('T')) {
    s += 'T00:00:00';
  }
  const lower = s.toLowerCase();
  if (!lower.endsWith('z') && !/[+-]\d{2}:?\d{2}$/.test(s)) {
    s += 'Z';
  }
  return s;
}

/**
 * Sanitise an attachment-provided file name. Strips path separators, control
 * chars, and leading dots. Returns `'attachment'` if everything's stripped.
 */
export function sanitiseAttachmentName(name: string | null | undefined): string {
  if (!name) return 'attachment';
  // Reject path components entirely.
  let s = name.split(/[\\/]/).pop() ?? '';
  s = s.replace(/[\x00-\x1f\x7f]/g, '');
  s = s.replace(/^\.+/, '');
  return s || 'attachment';
}

interface RawPageResponse<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

function envelopeOf<T>(raw: RawPageResponse<T>): PageEnvelope<T> {
  const results = raw.value ?? [];
  return {
    results,
    count: results.length,
    nextLink: raw['@odata.nextLink'] ?? null,
  };
}

export class MailResource {
  constructor(private readonly graph: Client) {}

  /** GET /me/mailFolders/<folder>/messages — list messages in a folder. */
  async list(options: ListMessagesOptions = {}): Promise<PageEnvelope<MessageSummary>> {
    if (options.focused && options.other) {
      throw new Error('--focused and --other are mutually exclusive.');
    }
    const limit = options.limit ?? 10;
    const folder = options.folder ?? 'inbox';

    const filters: string[] = [];
    if (options.unread) filters.push('isRead eq false');
    if (options.from) {
      const escaped = options.from.replace(/'/g, "''");
      filters.push(`from/emailAddress/address eq '${escaped}'`);
    }
    if (options.after) {
      filters.push(`receivedDateTime ge ${normaliseDateForFilter(options.after)}`);
    }
    if (options.before) {
      filters.push(`receivedDateTime le ${normaliseDateForFilter(options.before)}`);
    }
    if (options.focused) filters.push("inferenceClassification eq 'focused'");
    if (options.other) filters.push("inferenceClassification eq 'other'");

    const query: Record<string, string | number> = {
      $top: limit,
      $orderby: 'receivedDateTime DESC',
      $select: MESSAGE_SELECT_FIELDS,
    };
    if (filters.length > 0) {
      query.$filter = filters.join(' and ');
    }

    try {
      const folderId = await this.resolveFolderId(folder);
      const path = `/me/mailFolders/${encodeURIComponent(folderId)}/messages`;
      const raw = (await this.graph.api(path).query(query).get()) as RawPageResponse<MessageSummary>;
      return envelopeOf(raw);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /** GET /me/messages/<id> — full message. */
  async get(messageId: string, options: ReadMessageOptions = {}): Promise<MessageFull> {
    try {
      let req = this.graph
        .api(`/me/messages/${encodeURIComponent(messageId)}`)
        .query({ $select: MESSAGE_FULL_SELECT_FIELDS });
      if (options.preferText) {
        req = req.header('Prefer', 'outlook.body-content-type="text"');
      }
      return (await req.get()) as MessageFull;
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /** GET /me/messages?$search="..." — KQL search across all folders. */
  async search(options: SearchMessagesOptions): Promise<PageEnvelope<MessageSummary>> {
    const limit = options.limit ?? 25;
    // Graph requires $search values to be double-quoted.
    const searchValue = `"${options.query.replace(/"/g, '\\"')}"`;
    try {
      const raw = (await this.graph
        .api('/me/messages')
        .query({
          $top: limit,
          $search: searchValue,
          $select: MESSAGE_SELECT_FIELDS,
        })
        .get()) as RawPageResponse<MessageSummary>;
      return envelopeOf(raw);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /** GET /me/mailFolders — list folders with item counts. */
  async listFolders(): Promise<PageEnvelope<MailFolderSummary>> {
    try {
      const raw = (await this.graph
        .api('/me/mailFolders')
        .query({ $select: FOLDER_SELECT_FIELDS })
        .get()) as RawPageResponse<MailFolderSummary>;
      return envelopeOf(raw);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /** GET /me/messages/<id>/attachments — metadata only (no contentBytes). */
  async listAttachments(messageId: string): Promise<PageEnvelope<AttachmentSummary>> {
    try {
      const raw = (await this.graph
        .api(`/me/messages/${encodeURIComponent(messageId)}/attachments`)
        .query({ $select: ATTACHMENT_SELECT_FIELDS })
        .get()) as RawPageResponse<AttachmentSummary>;
      return envelopeOf(raw);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * GET /me/messages/<id>/attachments/<aid> — full FileAttachment with
   * contentBytes. Decodes the base64 payload and returns raw bytes; the caller
   * is responsible for writing them to disk.
   *
   * Throws CoreError for non-file attachment kinds (item / reference).
   */
  async downloadAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<DownloadAttachmentResult> {
    let raw: (FileAttachment & { '@odata.type'?: string }) | undefined;
    try {
      raw = (await this.graph
        .api(
          `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
        )
        .get()) as FileAttachment & { '@odata.type'?: string };
    } catch (err) {
      throw liftGraphError(err);
    }
    if (!raw) {
      throw new NotFoundError('attachment not found');
    }
    const odataType = raw['@odata.type'] ?? '';
    if (odataType && odataType !== '#microsoft.graph.fileAttachment') {
      throw new Error(
        `Cannot download attachment of kind ${odataType}; only fileAttachments expose contentBytes.`,
      );
    }
    const base64 = raw.contentBytes;
    if (typeof base64 !== 'string' || base64.length === 0) {
      throw new Error('Attachment has no content bytes.');
    }
    const contentBytes = Buffer.from(base64, 'base64');
    return {
      contentBytes,
      name: sanitiseAttachmentName(raw.name ?? null),
      contentType: raw.contentType ?? null,
      size: contentBytes.byteLength,
    };
  }

  /**
   * Resolve `--folder` flag → Graph folder id.
   *
   * - Well-known names (case-insensitive) → lowercase alias Graph accepts.
   * - Long opaque strings → assumed to already be ids.
   * - Anything else → look up displayName via /me/mailFolders.
   */
  private async resolveFolderId(folder: string): Promise<string> {
    const key = folder.toLowerCase();
    if (WELL_KNOWN_FOLDERS.has(key)) return key;
    if (looksLikeFolderId(folder)) return folder;

    const escaped = folder.replace(/'/g, "''");
    let raw: RawPageResponse<MailFolderSummary>;
    try {
      raw = (await this.graph
        .api('/me/mailFolders')
        .query({ $filter: `displayName eq '${escaped}'`, $select: 'id,displayName' })
        .get()) as RawPageResponse<MailFolderSummary>;
    } catch (err) {
      throw liftGraphError(err);
    }
    const first = raw.value?.[0];
    if (!first || !first.id) {
      throw new NotFoundError(`Folder '${folder}' not found.`);
    }
    return first.id;
  }
}
