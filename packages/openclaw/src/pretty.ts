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

/**
 * Common return shape for triage tools (move/delete/mark/flag/importance):
 * always carries an `id`, plus exactly one of the discriminator fields below.
 */
interface TriageResultShape {
  id: string;
  oldId?: string;
  destinationFolder?: string;
  deleted?: true;
  isRead?: boolean;
  flagStatus?: 'flagged' | 'complete' | 'notFlagged';
  importance?: 'low' | 'normal' | 'high';
}

interface AttendeeShape {
  address: string | null;
  name: string | null;
  type: string | null;
  response: string | null;
}

interface EventSummaryShape {
  id: string | null;
  subject: string | null;
  start: string | null;
  end: string | null;
  timeZone: string | null;
  location: string | null;
  organizer: string | null;
  attendees: AttendeeShape[];
  isOnlineMeeting: boolean | null;
  onlineJoinUrl: string | null;
  isAllDay: boolean | null;
  isCancelled: boolean | null;
  webLink: string | null;
}

interface EventDetailShape extends EventSummaryShape {
  bodyContentType: string | null;
  body: string | null;
  bodyPreview: string | null;
}

interface EventListShape {
  events: EventSummaryShape[];
  count: number;
  nextLink: string | null;
}

interface AvailabilityScheduleShape {
  scheduleId: string | null;
  availabilityView: string | null;
  scheduleItems: Array<{
    subject: string | null;
    start: string | null;
    end: string | null;
    status: string | null;
    location: string | null;
  }>;
  workingHours: unknown;
}

interface AvailabilityResultShape {
  emails: string[];
  startTime: string;
  endTime: string;
  timeZone: string;
  interval: number;
  schedules: AvailabilityScheduleShape[];
}

interface CalendarResponseShape {
  id: string;
  response: 'accept' | 'decline' | 'tentative';
  sentResponse: boolean;
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

function isEventList(p: unknown): p is EventListShape {
  return isObject(p) && Array.isArray((p as { events?: unknown }).events);
}

function isEventDetail(p: unknown): p is EventDetailShape {
  if (!isObject(p)) return false;
  return (
    'attendees' in p &&
    Array.isArray((p as { attendees: unknown }).attendees) &&
    'subject' in p &&
    'start' in p &&
    'body' in p
  );
}

function isEventSummary(p: unknown): p is EventSummaryShape {
  if (!isObject(p)) return false;
  return (
    'attendees' in p &&
    Array.isArray((p as { attendees: unknown }).attendees) &&
    'start' in p &&
    'end' in p &&
    'organizer' in p
  );
}

function isAvailabilityResult(p: unknown): p is AvailabilityResultShape {
  if (!isObject(p)) return false;
  return (
    Array.isArray((p as { schedules?: unknown }).schedules) &&
    Array.isArray((p as { emails?: unknown }).emails) &&
    typeof (p as { interval?: unknown }).interval === 'number'
  );
}

function isCalendarResponse(p: unknown): p is CalendarResponseShape {
  if (!isObject(p)) return false;
  return (
    typeof (p as { id?: unknown }).id === 'string' &&
    typeof (p as { sentResponse?: unknown }).sentResponse === 'boolean' &&
    'response' in p
  );
}

function isTriageResult(p: unknown): p is TriageResultShape {
  if (!isObject(p)) return false;
  if (typeof (p as { id?: unknown }).id !== 'string') return false;
  return (
    'deleted' in p ||
    'destinationFolder' in p ||
    'isRead' in p ||
    'flagStatus' in p ||
    'importance' in p
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
  if (isEventList(payload)) return renderEventList(payload);
  if (isEventDetail(payload)) return renderEventDetail(payload);
  if (isEventSummary(payload)) return renderEventDetail(payload as EventDetailShape);
  if (isAvailabilityResult(payload)) return renderAvailability(payload);
  if (isCalendarResponse(payload)) return renderCalendarResponse(payload);
  if (isTriageResult(payload)) return renderTriageResult(payload);

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

function renderTriageResult(p: TriageResultShape): string {
  if (p.deleted) {
    return `Deleted ${p.id} (moved to Deleted Items).`;
  }
  if (p.destinationFolder !== undefined) {
    const fromPart = p.oldId ? `${p.oldId} -> ` : '';
    return `Moved ${fromPart}${p.destinationFolder}\n  new id: ${p.id}`;
  }
  if (p.isRead !== undefined) {
    return `Marked ${p.isRead ? 'read' : 'unread'}: ${p.id}`;
  }
  if (p.flagStatus !== undefined) {
    return `Flag set ${p.flagStatus}: ${p.id}`;
  }
  if (p.importance !== undefined) {
    return `Importance set ${p.importance}: ${p.id}`;
  }
  // Should be unreachable given isTriageResult, but fall through safely.
  return `Updated ${p.id}.`;
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

function renderEventList(p: EventListShape): string {
  const lines: string[] = [`Events (${p.count})`];
  if (p.count === 0) {
    lines.push('  (none)');
  } else {
    for (const e of p.events) {
      const start = shortDate(e.start);
      const subject = e.subject ?? '(no subject)';
      const organizer = e.organizer ? ` (${e.organizer})` : '';
      const location = e.location ? ` [${e.location}]` : '';
      lines.push(`  ${start}  ${subject}${organizer}${location}`);
      if (e.id) lines.push(`    id: ${e.id}`);
    }
  }
  if (p.nextLink) lines.push(`  (more results — nextLink: ${p.nextLink})`);
  return lines.join('\n');
}

function renderEventDetail(e: EventDetailShape): string {
  const lines: string[] = [];
  lines.push(`Subject: ${e.subject ?? '(no subject)'}`);
  if (e.start) {
    const tz = e.timeZone ? ` (${e.timeZone})` : '';
    lines.push(`Start:   ${e.start}${tz}`);
  }
  if (e.end) lines.push(`End:     ${e.end}`);
  if (e.location) lines.push(`Where:   ${e.location}`);
  if (e.organizer) lines.push(`Organizer: ${e.organizer}`);
  if (e.attendees.length > 0) {
    lines.push('Attendees:');
    for (const a of e.attendees) {
      const resp = a.response ? ` (${a.response})` : '';
      lines.push(`  - ${a.address ?? '(no address)'}${resp}`);
    }
  }
  if (e.onlineJoinUrl) lines.push(`Join:    ${e.onlineJoinUrl}`);
  const body = e.body ?? '';
  if (body.length > 0) {
    lines.push('');
    if (body.length > 2000) {
      lines.push(body.slice(0, 2000));
      lines.push('…(body truncated; use output: json for full text)');
    } else {
      lines.push(body);
    }
  }
  if (e.webLink) {
    lines.push('');
    lines.push(`Open in Outlook: ${e.webLink}`);
  }
  return lines.join('\n');
}

function renderAvailability(r: AvailabilityResultShape): string {
  const lines: string[] = [];
  lines.push(`Availability  ${r.startTime} -> ${r.endTime}  ${r.timeZone}  (${r.interval}min blocks)`);
  lines.push('Legend: 0=free, 1=tentative, 2=busy, 3=out-of-office, 4=working-elsewhere');
  if (r.schedules.length === 0) {
    lines.push('  (no schedules returned)');
    return lines.join('\n');
  }
  for (const s of r.schedules) {
    lines.push(`  ${s.scheduleId ?? '(unknown)'}  ${s.availabilityView ?? ''}`);
    for (const item of s.scheduleItems) {
      const span =
        item.start && item.end
          ? `${shortDate(item.start)} -> ${shortDate(item.end)}`
          : '';
      const subject = item.subject ?? '';
      const status = item.status ? ` [${item.status}]` : '';
      const loc = item.location ? ` @ ${item.location}` : '';
      if (span || subject) {
        lines.push(`      ${span}  ${subject}${status}${loc}`);
      }
    }
  }
  return lines.join('\n');
}

function renderCalendarResponse(p: CalendarResponseShape): string {
  const sent = p.sentResponse ? '' : ' (no notification sent)';
  return `Responded ${p.response} to event ${p.id}${sent}.`;
}
