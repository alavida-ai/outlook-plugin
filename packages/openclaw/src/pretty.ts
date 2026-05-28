/**
 * Shape-detected pretty renderer for outlook tool results.
 *
 * Each renderer is a small `if (isMessageList(p)) return renderX(p);` branch
 * that fires when the result matches a known shape. Unknown shapes fall
 * through to a truncated-JSON dump.
 */
import { isToolErrorEnvelope, type ToolErrorEnvelope } from './errors.js';

interface MessageSummaryShape {
  id: string | null;
  subject: string | null;
  from: string | null;
  receivedDateTime: string | null;
  isRead: boolean | null;
  hasAttachments: boolean | null;
  bodyPreview: string | null;
  webLink: string | null;
}

interface MessageFullShape extends MessageSummaryShape {
  to: string[];
  cc: string[];
  bcc: string[];
  importance: string | null;
  bodyContentType: string | null;
  body: string | null;
}

interface MessageListShape {
  messages: MessageSummaryShape[];
  count: number;
  nextLink: string | null;
}

interface FolderListShape {
  folders: Array<{
    id: string | null;
    displayName: string | null;
    unreadItemCount: number;
    totalItemCount: number;
  }>;
  count: number;
  nextLink: string | null;
}

interface AttachmentListShape {
  attachments: Array<{
    id: string | null;
    name: string | null;
    contentType: string | null;
    size: number | null;
    isInline: boolean;
  }>;
  count: number;
  nextLink: string | null;
}

interface DownloadResultShape {
  path: string;
  name: string;
  contentType: string | null;
  size: number;
}

interface DraftSummaryShape {
  id: string;
  subject: string | null;
  to: string[];
  cc: string[];
  bcc: string[];
  webLink: string | null;
  composeLink: string | null;
}

interface AddAttachmentSummaryShape {
  attachmentId: string | null;
  name: string | null;
  contentType: string | null;
  size: number | null;
  isInline: boolean;
  draftId?: string;
}

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

function isMessageList(p: unknown): p is MessageListShape {
  return isObject(p) && Array.isArray((p as { messages?: unknown }).messages);
}

function isFolderList(p: unknown): p is FolderListShape {
  return isObject(p) && Array.isArray((p as { folders?: unknown }).folders);
}

function isAttachmentList(p: unknown): p is AttachmentListShape {
  return isObject(p) && Array.isArray((p as { attachments?: unknown }).attachments);
}

function isMessageFull(p: unknown): p is MessageFullShape {
  if (!isObject(p)) return false;
  return (
    'subject' in p &&
    'to' in p &&
    Array.isArray((p as { to: unknown }).to) &&
    'body' in p
  );
}

function isDownloadResult(p: unknown): p is DownloadResultShape {
  if (!isObject(p)) return false;
  return (
    typeof (p as { path?: unknown }).path === 'string' &&
    typeof (p as { name?: unknown }).name === 'string' &&
    typeof (p as { size?: unknown }).size === 'number'
  );
}

function isDraftSummary(p: unknown): p is DraftSummaryShape {
  if (!isObject(p)) return false;
  return (
    typeof (p as { id?: unknown }).id === 'string' &&
    'composeLink' in p &&
    'webLink' in p &&
    Array.isArray((p as { to?: unknown }).to)
  );
}

function isAddAttachmentSummary(p: unknown): p is AddAttachmentSummaryShape {
  if (!isObject(p)) return false;
  return (
    'attachmentId' in p &&
    'isInline' in p &&
    'size' in p &&
    !('path' in p)
  );
}

/** Render an arbitrary tool payload as compact text. */
export function renderPretty(payload: unknown): string {
  if (payload === undefined || payload === null) return '(no result)';
  if (isToolErrorEnvelope(payload)) return renderError(payload);
  if (isMessageList(payload)) return renderMessageList(payload);
  if (isMessageFull(payload)) return renderMessageFull(payload);
  if (isFolderList(payload)) return renderFolderList(payload);
  if (isAttachmentList(payload)) return renderAttachmentList(payload);
  if (isDownloadResult(payload)) return renderDownloadResult(payload);
  if (isAddAttachmentSummary(payload)) return renderAddAttachmentSummary(payload);
  if (isDraftSummary(payload)) return renderDraftSummary(payload);

  // Generic fallback — JSON.stringify (truncated for readability).
  try {
    const text = JSON.stringify(payload, null, 2);
    return text.length > 4000 ? text.slice(0, 4000) + '\n…(truncated; use output: json)' : text;
  } catch {
    return String(payload);
  }
}

function renderError(envelope: ToolErrorEnvelope): string {
  const e = envelope.__toolError;
  const lines = [`✗ ${e.error}`, `  ${e.message}`];
  if (e.hint) lines.push(`  → ${e.hint}`);
  if (e.retryAfterSeconds !== undefined) lines.push(`  retry after: ${e.retryAfterSeconds}s`);
  if (e.accounts && e.accounts.length > 0) {
    lines.push(`  accounts:    ${e.accounts.join(', ')}`);
  }
  return lines.join('\n');
}

function shortDate(s: string | null): string {
  if (!s) return '';
  return s.replace('T', ' ').slice(0, 16);
}

function renderMessageList(p: MessageListShape): string {
  const lines: string[] = [`Messages (${p.count})`];
  if (p.count === 0) {
    lines.push('  (none)');
  } else {
    for (const m of p.messages) {
      const unread = m.isRead === false ? '* ' : '  ';
      const att = m.hasAttachments ? ' [att]' : '';
      lines.push(
        `${unread}${shortDate(m.receivedDateTime)}  ${m.from ?? ''}  ${m.subject ?? '(no subject)'}${att}`,
      );
      if (m.id) lines.push(`    id: ${m.id}`);
    }
  }
  if (p.nextLink) lines.push(`  (more results — nextLink: ${p.nextLink})`);
  return lines.join('\n');
}

function renderMessageFull(m: MessageFullShape): string {
  const lines: string[] = [];
  lines.push(`Subject: ${m.subject ?? '(no subject)'}`);
  lines.push(`From:    ${m.from ?? '(unknown)'}`);
  if (m.to.length > 0) lines.push(`To:      ${m.to.join(', ')}`);
  if (m.cc.length > 0) lines.push(`Cc:      ${m.cc.join(', ')}`);
  if (m.receivedDateTime) lines.push(`Date:    ${m.receivedDateTime}`);
  if (m.importance && m.importance !== 'normal') {
    lines.push(`Importance: ${m.importance}`);
  }
  lines.push('');
  const body = m.body ?? '';
  if (body.length > 4000) {
    lines.push(body.slice(0, 4000));
    lines.push('…(body truncated; use output: json for full text)');
  } else {
    lines.push(body);
  }
  if (m.webLink) {
    lines.push('');
    lines.push(`Open in Outlook: ${m.webLink}`);
  }
  return lines.join('\n');
}

function renderFolderList(p: FolderListShape): string {
  const lines: string[] = [`Mail folders (${p.count})`];
  if (p.count === 0) {
    lines.push('  (none)');
  } else {
    for (const f of p.folders) {
      lines.push(
        `  ${f.displayName ?? '(unnamed)'}  (${f.unreadItemCount}/${f.totalItemCount})  ${f.id ?? ''}`,
      );
    }
  }
  return lines.join('\n');
}

function renderAttachmentList(p: AttachmentListShape): string {
  const lines: string[] = [`Attachments (${p.count})`];
  if (p.count === 0) {
    lines.push('  (none)');
  } else {
    for (const a of p.attachments) {
      const inline = a.isInline ? ' (inline)' : '';
      lines.push(
        `  ${a.name ?? '(no-name)'}  ${a.size ?? 0}B  ${a.contentType ?? ''}${inline}`,
      );
      if (a.id) lines.push(`    id: ${a.id}`);
    }
  }
  return lines.join('\n');
}

function renderDownloadResult(p: DownloadResultShape): string {
  return `Wrote ${p.name} (${p.size} bytes) to ${p.path}`;
}

function renderDraftSummary(p: DraftSummaryShape): string {
  const lines: string[] = [];
  const subjectLabel = p.subject ?? '(no subject)';
  lines.push(`Draft "${subjectLabel}" created.`);
  lines.push(`  id: ${p.id}`);
  if (p.to.length > 0) lines.push(`  to: ${p.to.join(', ')}`);
  if (p.cc.length > 0) lines.push(`  cc: ${p.cc.join(', ')}`);
  if (p.bcc.length > 0) lines.push(`  bcc: ${p.bcc.join(', ')}`);
  if (p.composeLink) {
    lines.push(`  edit in Outlook: ${p.composeLink}`);
  } else if (p.webLink) {
    lines.push(`  open in Outlook: ${p.webLink}`);
  }
  return lines.join('\n');
}

function renderAddAttachmentSummary(p: AddAttachmentSummaryShape): string {
  const size = p.size ?? 0;
  const name = p.name ?? '(unnamed)';
  const draftPart = p.draftId ? ` to draft ${p.draftId}` : '';
  const lines = [`Attached ${name} (${size} bytes)${draftPart}.`];
  if (p.attachmentId) lines.push(`  attachmentId: ${p.attachmentId}`);
  if (p.contentType) lines.push(`  contentType: ${p.contentType}`);
  if (p.isInline) lines.push(`  inline: true`);
  return lines.join('\n');
}
