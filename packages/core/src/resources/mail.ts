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

/** Input for `MailResource.draft`. */
export interface DraftInput {
  subject: string;
  body: string;
  /** When true, body is sent as HTML; otherwise plain text. */
  html?: boolean;
  to: string[];
  cc?: string[];
  bcc?: string[];
}

/** Input for `MailResource.reply`. */
export interface ReplyInput {
  body: string;
  html?: boolean;
  /** Reply to all recipients on the thread instead of just the sender. */
  replyAll?: boolean;
}

/** Input for `MailResource.forward`. */
export interface ForwardInput {
  to: string[];
  cc?: string[];
  /** Optional leading note prepended above the quoted original. */
  comment?: string;
}

/** Input for `MailResource.addAttachment`. */
export interface AddAttachmentInput {
  /** File name as it should appear on the attachment. */
  name: string;
  /** MIME type (`application/pdf`, `image/png`, ...). */
  contentType: string;
  /** Raw bytes; will be base64-encoded for Graph. */
  contentBytes: Buffer;
  /** Display inline in the body instead of as a download. */
  isInline?: boolean;
}

/** Returned by `draft` / `reply` / `forward`. */
export interface DraftSummary {
  id: string;
  subject: string | null;
  to: string[];
  cc: string[];
  bcc: string[];
  webLink: string | null;
  /** outlook.cloud.microsoft compose URL — opens the draft in edit mode. */
  composeLink: string | null;
}

/** Returned by `addAttachment`. */
export interface AddAttachmentSummary {
  attachmentId: string | null;
  name: string | null;
  contentType: string | null;
  size: number | null;
  isInline: boolean;
}

/** Inline-attachment cap per Graph docs. Anything larger needs an upload session. */
export const INLINE_ATTACHMENT_MAX_BYTES = 3 * 1024 * 1024;

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

/** Wrap an address into Graph's nested Recipient shape. */
function toRecipient(address: string): { emailAddress: { address: string } } {
  return { emailAddress: { address } };
}

function toRecipients(
  addresses: readonly string[] | undefined,
): Array<{ emailAddress: { address: string } }> {
  if (!addresses || addresses.length === 0) return [];
  return addresses.map(toRecipient);
}

function recipientAddresses(
  list: Array<{ emailAddress?: { address?: string | null } | null }> | null | undefined,
): string[] {
  if (!list) return [];
  const out: string[] = [];
  for (const r of list) {
    const a = r.emailAddress?.address;
    if (a) out.push(a);
  }
  return out;
}

/**
 * Build the `outlook.cloud.microsoft/mail/compose/<itemId>` URL by extracting
 * the `ItemID` query param from Graph's `webLink`. Graph's webLink itself
 * opens the message in read mode; the compose route opens it in edit mode.
 *
 * Returns `null` if the webLink is missing or doesn't carry an ItemID.
 */
export function composeLinkFromWebLink(webLink: string | null | undefined): string | null {
  if (!webLink) return null;
  let parsed: URL;
  try {
    parsed = new URL(webLink);
  } catch {
    return null;
  }
  const itemId = parsed.searchParams.get('ItemID');
  if (!itemId) return null;
  return `https://outlook.cloud.microsoft/mail/compose/${encodeURIComponent(itemId)}`;
}

function draftSummaryFrom(msg: Message | null | undefined): DraftSummary {
  if (!msg || !msg.id) {
    throw new Error('Graph returned a draft response without an id.');
  }
  return {
    id: msg.id,
    subject: msg.subject ?? null,
    to: recipientAddresses(msg.toRecipients ?? null),
    cc: recipientAddresses(msg.ccRecipients ?? null),
    bcc: recipientAddresses(msg.bccRecipients ?? null),
    webLink: msg.webLink ?? null,
    composeLink: composeLinkFromWebLink(msg.webLink ?? null),
  };
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
   * POST /me/messages — create a draft. Graph drops it into the user's
   * Drafts folder; nothing is sent. The plugin never sends mail by design.
   */
  async draft(input: DraftInput): Promise<DraftSummary> {
    if (input.to.length === 0) {
      throw new Error('At least one --to recipient is required.');
    }
    const payload = {
      subject: input.subject,
      body: {
        contentType: input.html ? 'HTML' : 'Text',
        content: input.body,
      },
      toRecipients: toRecipients(input.to),
      ccRecipients: toRecipients(input.cc),
      bccRecipients: toRecipients(input.bcc),
    };
    try {
      const created = (await this.graph.api('/me/messages').post(payload)) as Message;
      return draftSummaryFrom(created);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * POST /me/messages/<id>/createReply (or createReplyAll) followed by PATCH
   * of the body. Graph's createReply seeds the draft with the quoted
   * original; the PATCH overlays the user's reply text on top.
   */
  async reply(messageId: string, input: ReplyInput): Promise<DraftSummary> {
    const endpoint = input.replyAll ? 'createReplyAll' : 'createReply';
    const idEnc = encodeURIComponent(messageId);
    try {
      // Step 1: createReply / createReplyAll. POST with no body returns the
      // seeded draft message resource.
      const seeded = (await this.graph
        .api(`/me/messages/${idEnc}/${endpoint}`)
        .post({})) as Message;
      if (!seeded || !seeded.id) {
        throw new Error('Graph did not return a draft id from createReply.');
      }
      // Step 2: PATCH the body onto the seeded draft. Graph preserves the
      // quoted original below the new content.
      const patched = (await this.graph
        .api(`/me/messages/${encodeURIComponent(seeded.id)}`)
        .patch({
          body: {
            contentType: input.html ? 'HTML' : 'Text',
            content: input.body,
          },
        })) as Message;
      // PATCH may return a thin object; fall back to the seeded message for
      // any field PATCH dropped (e.g. recipients, webLink).
      return draftSummaryFrom({ ...seeded, ...patched });
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * POST /me/messages/<id>/createForward with a comment and recipients.
   * Graph composes the forward draft with the quoted original; the comment
   * is prepended above it.
   */
  async forward(messageId: string, input: ForwardInput): Promise<DraftSummary> {
    if (input.to.length === 0) {
      throw new Error('At least one --to recipient is required for forward.');
    }
    const idEnc = encodeURIComponent(messageId);
    const payload: Record<string, unknown> = {
      comment: input.comment ?? '',
      toRecipients: toRecipients(input.to),
    };
    if (input.cc && input.cc.length > 0) {
      payload.ccRecipients = toRecipients(input.cc);
    }
    try {
      const created = (await this.graph
        .api(`/me/messages/${idEnc}/createForward`)
        .post(payload)) as Message;
      return draftSummaryFrom(created);
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * POST /me/messages/<draftId>/attachments — attach a file inline.
   *
   * Caps at 3 MB (Graph's inline-attachment limit). Larger files require
   * the createUploadSession + chunked PUT flow, which is intentionally out
   * of scope for this slice.
   */
  async addAttachment(
    draftId: string,
    input: AddAttachmentInput,
  ): Promise<AddAttachmentSummary> {
    if (input.contentBytes.byteLength > INLINE_ATTACHMENT_MAX_BYTES) {
      throw new Error(
        `Attachment is ${input.contentBytes.byteLength} bytes; inline attachments cap at ${INLINE_ATTACHMENT_MAX_BYTES} bytes (3 MB). Files this large require Graph's createUploadSession API, which is not yet implemented.`,
      );
    }
    const payload = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: input.name,
      contentType: input.contentType,
      contentBytes: input.contentBytes.toString('base64'),
      isInline: input.isInline ?? false,
    };
    try {
      const created = (await this.graph
        .api(`/me/messages/${encodeURIComponent(draftId)}/attachments`)
        .post(payload)) as FileAttachment;
      return {
        attachmentId: created?.id ?? null,
        name: created?.name ?? input.name,
        contentType: created?.contentType ?? input.contentType,
        size: created?.size ?? input.contentBytes.byteLength,
        isInline: created?.isInline ?? (input.isInline ?? false),
      };
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * POST /me/messages/<id>/move — move a message to another folder.
   *
   * Outlook reassigns the message id when it moves, so the returned `id` is
   * different from the input `messageId`. `destinationFolder` reports the
   * resolved id/well-known name we actually sent to Graph.
   */
  async move(
    messageId: string,
    folder: string,
  ): Promise<{ id: string; oldId: string; destinationFolder: string }> {
    try {
      const destinationFolder = await this.resolveFolderId(folder);
      const moved = (await this.graph
        .api(`/me/messages/${encodeURIComponent(messageId)}/move`)
        .post({ destinationId: destinationFolder })) as Message;
      if (!moved || !moved.id) {
        throw new Error('Graph returned a move response without an id.');
      }
      return { id: moved.id, oldId: messageId, destinationFolder };
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /**
   * DELETE /me/messages/<id> — soft-delete (Outlook moves to Deleted Items).
   *
   * Recoverable: the message lives in Deleted Items until the user empties
   * the folder. There is no hard-delete in this surface.
   */
  async delete(messageId: string): Promise<{ id: string; deleted: true }> {
    try {
      await this.graph.api(`/me/messages/${encodeURIComponent(messageId)}`).delete();
      return { id: messageId, deleted: true };
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /** PATCH /me/messages/<id> — set the read/unread state. */
  async mark(
    messageId: string,
    isRead: boolean,
  ): Promise<{ id: string; isRead: boolean }> {
    try {
      await this.graph
        .api(`/me/messages/${encodeURIComponent(messageId)}`)
        .patch({ isRead });
      return { id: messageId, isRead };
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /** PATCH /me/messages/<id> — set the follow-up flag status. */
  async flag(
    messageId: string,
    status: 'flagged' | 'complete' | 'notFlagged',
  ): Promise<{ id: string; flagStatus: 'flagged' | 'complete' | 'notFlagged' }> {
    try {
      await this.graph
        .api(`/me/messages/${encodeURIComponent(messageId)}`)
        .patch({ flag: { flagStatus: status } });
      return { id: messageId, flagStatus: status };
    } catch (err) {
      throw liftGraphError(err);
    }
  }

  /** PATCH /me/messages/<id> — set the importance level. */
  async importance(
    messageId: string,
    level: 'low' | 'normal' | 'high',
  ): Promise<{ id: string; importance: 'low' | 'normal' | 'high' }> {
    try {
      await this.graph
        .api(`/me/messages/${encodeURIComponent(messageId)}`)
        .patch({ importance: level });
      return { id: messageId, importance: level };
    } catch (err) {
      throw liftGraphError(err);
    }
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
