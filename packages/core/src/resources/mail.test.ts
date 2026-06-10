import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Client } from '@microsoft/microsoft-graph-client';

import { OutlookClient } from '../client.js';
import {
  composeLinkFromWebLink,
  inboxLinkFromId,
  normaliseDateForFilter,
  sanitiseAttachmentName,
} from './mail.js';

/**
 * Stand-in for the cheap `?$select=id,receivedDateTime,isDraft` pre-flight
 * the resource fires before listAttachments / downloadAttachment /
 * reply / forward to enforce the inbound-mail age filter. Use this when a
 * test doesn't care about the freshness check itself — it just needs the
 * pre-flight to pass so the operation under test can run.
 */
const PRE_FLIGHT_OLD_INBOUND = {
  id: 'pre-flight',
  receivedDateTime: '2020-01-01T00:00:00Z',
  isDraft: false,
} as const;

interface FakeRequest {
  api: ReturnType<typeof vi.fn>;
  query: ReturnType<typeof vi.fn>;
  header: ReturnType<typeof vi.fn>;
  get: ReturnType<typeof vi.fn>;
  post: ReturnType<typeof vi.fn>;
  patch: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
  capturedPaths: string[];
  capturedQueries: Array<Record<string, unknown>>;
  capturedHeaders: Array<[string, string]>;
  capturedPosts: unknown[];
  capturedPatches: unknown[];
  capturedMethods: string[];
}

/**
 * Build a fake Graph client that records the .api() path, all .query() args
 * and .header() args, and returns the supplied responses in order.
 *
 * Each call to .api() returns a chainable object whose .get()/.post()/.patch()
 * resolves to the next queued response. A single shared queue is used so the
 * caller controls the response order independent of which HTTP verb fires.
 */
function fakeGraph(responses: unknown[]): { graph: Client; calls: FakeRequest } {
  const queue = [...responses];
  const calls: FakeRequest = {
    api: vi.fn(),
    query: vi.fn(),
    header: vi.fn(),
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
    capturedPaths: [],
    capturedQueries: [],
    capturedHeaders: [],
    capturedPosts: [],
    capturedPatches: [],
    capturedMethods: [],
  };

  const chain = {
    query(q: Record<string, unknown>) {
      calls.capturedQueries.push(q);
      calls.query(q);
      return chain;
    },
    header(k: string, v: string) {
      calls.capturedHeaders.push([k, v]);
      calls.header(k, v);
      return chain;
    },
    async get() {
      calls.capturedMethods.push('GET');
      calls.get();
      return queue.shift();
    },
    async post(body: unknown) {
      calls.capturedMethods.push('POST');
      calls.capturedPosts.push(body);
      calls.post(body);
      return queue.shift();
    },
    async patch(body: unknown) {
      calls.capturedMethods.push('PATCH');
      calls.capturedPatches.push(body);
      calls.patch(body);
      return queue.shift();
    },
    async delete() {
      calls.capturedMethods.push('DELETE');
      calls.delete();
      return queue.shift();
    },
  };

  const graph = {
    api(path: string) {
      calls.capturedPaths.push(path);
      calls.api(path);
      return chain;
    },
  } as unknown as Client;

  return { graph, calls };
}

describe('normaliseDateForFilter', () => {
  it('expands a bare YYYY-MM-DD to UTC midnight Z', () => {
    expect(normaliseDateForFilter('2026-05-28')).toBe('2026-05-28T00:00:00Z');
  });
  it('passes through ISO8601 with Z', () => {
    expect(normaliseDateForFilter('2026-05-28T15:00:00Z')).toBe('2026-05-28T15:00:00Z');
  });
  it('passes through ISO8601 with offset', () => {
    expect(normaliseDateForFilter('2026-05-28T15:00:00+02:00')).toBe(
      '2026-05-28T15:00:00+02:00',
    );
  });
});

describe('sanitiseAttachmentName', () => {
  it('strips path separators', () => {
    expect(sanitiseAttachmentName('a/b/c.pdf')).toBe('c.pdf');
    expect(sanitiseAttachmentName('..\\..\\evil.exe')).toBe('evil.exe');
  });
  it('strips leading dots', () => {
    expect(sanitiseAttachmentName('.hidden')).toBe('hidden');
  });
  it('falls back to default when empty', () => {
    expect(sanitiseAttachmentName('')).toBe('attachment');
    expect(sanitiseAttachmentName(null)).toBe('attachment');
    expect(sanitiseAttachmentName('...')).toBe('attachment');
  });
  it('strips control chars', () => {
    expect(sanitiseAttachmentName('a\x00b\x1fc.txt')).toBe('abc.txt');
  });
});

describe('MailResource.list', () => {
  it('hits the well-known folder path with default query', async () => {
    const { graph, calls } = fakeGraph([
      { value: [{ id: 'm1', subject: 'hi' }] },
    ]);
    const out = new OutlookClient(graph);
    const page = await out.mail.list();
    expect(calls.capturedPaths).toEqual(['/me/mailFolders/inbox/messages']);
    expect(calls.capturedQueries[0]).toMatchObject({
      $top: 10,
      $orderby: 'receivedDateTime DESC',
    });
    expect(page.results).toHaveLength(1);
    expect(page.count).toBe(1);
    expect(page.nextLink).toBeNull();
  });

  it('composes filter clauses', async () => {
    // Pin the clock so the trailing quarantine cutoff is deterministic.
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-05-29T00:00:00Z'));
    try {
      const { graph, calls } = fakeGraph([{ value: [] }]);
      const out = new OutlookClient(graph);
      await out.mail.list({
        limit: 5,
        folder: 'sentitems',
        unread: true,
        from: 'boss@example.com',
        after: '2026-05-01',
        before: '2026-05-28',
        focused: true,
      });
      expect(calls.capturedPaths[0]).toBe('/me/mailFolders/sentitems/messages');
      const q = calls.capturedQueries[0] as Record<string, unknown>;
      expect(q.$top).toBe(5);
      // Trailing clause exempts drafts and filters anything received in the
      // last 30 minutes (the inbound-mail age filter).
      expect(q.$filter).toBe(
        "isRead eq false and from/emailAddress/address eq 'boss@example.com' and receivedDateTime ge 2026-05-01T00:00:00Z and receivedDateTime le 2026-05-28T00:00:00Z and inferenceClassification eq 'focused' and (isDraft eq true or receivedDateTime le 2026-05-28T23:30:00.000Z)",
      );
    } finally {
      vi.useRealTimers();
    }
  });

  it('resolves a custom folder displayName', async () => {
    const { graph, calls } = fakeGraph([
      { value: [{ id: 'fld-xyz', displayName: 'Projects' }] },
      { value: [{ id: 'm2', subject: 'p' }] },
    ]);
    const out = new OutlookClient(graph);
    const page = await out.mail.list({ folder: 'Projects' });
    expect(calls.capturedPaths).toEqual([
      '/me/mailFolders',
      '/me/mailFolders/fld-xyz/messages',
    ]);
    expect(page.results).toHaveLength(1);
  });

  it('rejects --focused + --other', async () => {
    const { graph } = fakeGraph([{ value: [] }]);
    const out = new OutlookClient(graph);
    await expect(out.mail.list({ focused: true, other: true })).rejects.toThrow(
      'mutually exclusive',
    );
  });

  it('propagates @odata.nextLink', async () => {
    const { graph } = fakeGraph([
      { value: [{ id: 'a' }], '@odata.nextLink': 'https://graph/next' },
    ]);
    const out = new OutlookClient(graph);
    const page = await out.mail.list();
    expect(page.nextLink).toBe('https://graph/next');
  });
});

describe('MailResource.get', () => {
  it('fetches /me/messages/<id> with select', async () => {
    const { graph, calls } = fakeGraph([{ id: 'm', subject: 'hi' }]);
    const out = new OutlookClient(graph);
    const msg = await out.mail.get('AAA-id');
    expect(calls.capturedPaths).toEqual(['/me/messages/AAA-id']);
    expect(calls.capturedHeaders).toEqual([]);
    expect(msg.subject).toBe('hi');
  });
  it('adds the Prefer header when preferText is set', async () => {
    const { graph, calls } = fakeGraph([{ id: 'm', subject: 'p' }]);
    const out = new OutlookClient(graph);
    await out.mail.get('m-id', { preferText: true });
    expect(calls.capturedHeaders).toEqual([
      ['Prefer', 'outlook.body-content-type="text"'],
    ]);
  });
  it('url-encodes message ids', async () => {
    const { graph, calls } = fakeGraph([{ id: 'm' }]);
    const out = new OutlookClient(graph);
    await out.mail.get('AAA/BBB=');
    expect(calls.capturedPaths[0]).toBe('/me/messages/AAA%2FBBB%3D');
  });
});

describe('inboxLinkFromId', () => {
  it('builds the cloud.microsoft inbox URL with the id', () => {
    expect(inboxLinkFromId('AAQkADg3NTE0NWVkLT')).toBe(
      'https://outlook.cloud.microsoft/mail/inbox/id/AAQkADg3NTE0NWVkLT',
    );
  });
  it('percent-encodes payload characters that need it', () => {
    expect(inboxLinkFromId('AAQk+/=')).toBe(
      'https://outlook.cloud.microsoft/mail/inbox/id/AAQk%2B%2F%3D',
    );
  });
  it('returns null when id is missing', () => {
    expect(inboxLinkFromId(null)).toBeNull();
    expect(inboxLinkFromId(undefined)).toBeNull();
    expect(inboxLinkFromId('')).toBeNull();
  });
});

describe('MailResource.inboxLinks', () => {
  it('POSTs translateExchangeIds with restId → restImmutableEntryId and builds URLs', async () => {
    const { graph, calls } = fakeGraph([
      {
        value: [
          { sourceId: 'AAMk-1', targetId: 'AAQk-1' },
          { sourceId: 'AAMk-2', targetId: 'AAQk-2+/' },
        ],
      },
    ]);
    const out = new OutlookClient(graph);
    const links = await out.mail.inboxLinks(['AAMk-1', 'AAMk-2']);
    expect(calls.capturedPaths).toEqual(['/me/translateExchangeIds']);
    expect(calls.capturedMethods).toEqual(['POST']);
    expect(calls.capturedPosts[0]).toEqual({
      inputIds: ['AAMk-1', 'AAMk-2'],
      sourceIdType: 'restId',
      targetIdType: 'restImmutableEntryId',
    });
    expect(links).toEqual({
      'AAMk-1': 'https://outlook.cloud.microsoft/mail/inbox/id/AAQk-1',
      'AAMk-2': 'https://outlook.cloud.microsoft/mail/inbox/id/AAQk-2%2B%2F',
    });
  });

  it('returns null for ids Graph failed to translate', async () => {
    const { graph } = fakeGraph([
      { value: [{ sourceId: 'AAMk-1', targetId: 'AAQk-1' }] },
    ]);
    const out = new OutlookClient(graph);
    const links = await out.mail.inboxLinks(['AAMk-1', 'AAMk-missing']);
    expect(links).toEqual({
      'AAMk-1': 'https://outlook.cloud.microsoft/mail/inbox/id/AAQk-1',
      'AAMk-missing': null,
    });
  });

  it('short-circuits with empty input', async () => {
    const { graph, calls } = fakeGraph([]);
    const out = new OutlookClient(graph);
    const links = await out.mail.inboxLinks([]);
    expect(links).toEqual({});
    expect(calls.capturedPaths).toEqual([]); // no API call
  });
});

describe('MailResource.search', () => {
  it('quotes the query and uses $search', async () => {
    const { graph, calls } = fakeGraph([{ value: [{ id: 'r1' }] }]);
    const out = new OutlookClient(graph);
    const page = await out.mail.search({ query: 'from:boss subject:urgent', limit: 3 });
    expect(calls.capturedPaths).toEqual(['/me/messages']);
    const q = calls.capturedQueries[0] as Record<string, unknown>;
    expect(q.$top).toBe(3);
    expect(q.$search).toBe('"from:boss subject:urgent"');
    expect(q.$orderby).toBeUndefined();
    expect(page.results).toHaveLength(1);
  });
});

describe('MailResource.listFolders', () => {
  it('hits /me/mailFolders with select', async () => {
    const { graph, calls } = fakeGraph([
      {
        value: [
          { id: 'f1', displayName: 'Inbox', unreadItemCount: 4, totalItemCount: 200 },
        ],
      },
    ]);
    const out = new OutlookClient(graph);
    const page = await out.mail.listFolders();
    expect(calls.capturedPaths).toEqual(['/me/mailFolders']);
    expect(calls.capturedQueries[0]).toMatchObject({
      $select: 'id,displayName,unreadItemCount,totalItemCount',
    });
    expect(page.results[0]?.displayName).toBe('Inbox');
  });
});

describe('MailResource.listAttachments', () => {
  it('hits /me/messages/<id>/attachments with select', async () => {
    const { graph, calls } = fakeGraph([
      PRE_FLIGHT_OLD_INBOUND,
      {
        value: [
          { id: 'a1', name: 'file.pdf', contentType: 'application/pdf', size: 100, isInline: false },
        ],
      },
    ]);
    const out = new OutlookClient(graph);
    const page = await out.mail.listAttachments('msg-1');
    expect(calls.capturedPaths).toEqual([
      '/me/messages/msg-1',
      '/me/messages/msg-1/attachments',
    ]);
    expect(page.results[0]?.name).toBe('file.pdf');
  });
});

describe('MailResource.downloadAttachment', () => {
  it('decodes contentBytes and returns metadata', async () => {
    const payload = Buffer.from('hello world');
    const { graph, calls } = fakeGraph([
      PRE_FLIGHT_OLD_INBOUND,
      {
        '@odata.type': '#microsoft.graph.fileAttachment',
        id: 'a1',
        name: 'hello.txt',
        contentType: 'text/plain',
        contentBytes: payload.toString('base64'),
      },
    ]);
    const out = new OutlookClient(graph);
    const result = await out.mail.downloadAttachment('msg-1', 'a1');
    expect(calls.capturedPaths).toEqual([
      '/me/messages/msg-1',
      '/me/messages/msg-1/attachments/a1',
    ]);
    expect(result.name).toBe('hello.txt');
    expect(result.contentType).toBe('text/plain');
    expect(result.size).toBe(payload.byteLength);
    expect(Buffer.compare(result.contentBytes, payload)).toBe(0);
  });
  it('rejects non-fileAttachment kinds', async () => {
    const { graph } = fakeGraph([
      PRE_FLIGHT_OLD_INBOUND,
      { '@odata.type': '#microsoft.graph.itemAttachment', id: 'a2', name: 'card' },
    ]);
    const out = new OutlookClient(graph);
    await expect(out.mail.downloadAttachment('m', 'a2')).rejects.toThrow(
      'Cannot download attachment',
    );
  });
  it('sanitises the returned name', async () => {
    const { graph } = fakeGraph([
      PRE_FLIGHT_OLD_INBOUND,
      {
        '@odata.type': '#microsoft.graph.fileAttachment',
        id: 'a3',
        name: '../etc/passwd',
        contentType: 'text/plain',
        contentBytes: Buffer.from('x').toString('base64'),
      },
    ]);
    const out = new OutlookClient(graph);
    const r = await out.mail.downloadAttachment('m', 'a3');
    expect(r.name).toBe('passwd');
  });
});

describe('MailResource — inbound-mail age filter', () => {
  // Pin the clock so the cutoff is deterministic across every test in this
  // block. Cutoff = system time − 30 min.
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-05-29T12:00:00Z'));
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  // 30-min cutoff at 12:00Z → 11:30Z.
  const FRESH_RECEIVED = '2026-05-29T11:50:00Z'; // 10 min ago — quarantined
  const OLD_RECEIVED = '2026-05-29T10:00:00Z'; //  2 h ago — allowed

  it('list: drafts are exempt; non-draft fresh mail is filtered server-side', async () => {
    const { graph, calls } = fakeGraph([{ value: [] }]);
    const out = new OutlookClient(graph);
    await out.mail.list();
    const q = calls.capturedQueries[0] as Record<string, unknown>;
    expect(q.$filter).toBe(
      '(isDraft eq true or receivedDateTime le 2026-05-29T11:30:00.000Z)',
    );
  });

  it('get: refuses fresh non-draft messages with MailQuarantinedError', async () => {
    const { graph } = fakeGraph([
      { id: 'msg-fresh', isDraft: false, receivedDateTime: FRESH_RECEIVED, body: { content: 'OTP: 123456' } },
    ]);
    const out = new OutlookClient(graph);
    await expect(out.mail.get('msg-fresh')).rejects.toThrow(/safety window/i);
  });

  it('get: allows fresh drafts (the user is composing them)', async () => {
    const { graph } = fakeGraph([
      { id: 'd1', isDraft: true, receivedDateTime: FRESH_RECEIVED, body: { content: 'my draft' } },
    ]);
    const out = new OutlookClient(graph);
    const msg = await out.mail.get('d1');
    expect(msg.body?.content).toBe('my draft');
  });

  it('get: allows older non-draft messages', async () => {
    const { graph } = fakeGraph([
      { id: 'm1', isDraft: false, receivedDateTime: OLD_RECEIVED, body: { content: 'old mail' } },
    ]);
    const out = new OutlookClient(graph);
    const msg = await out.mail.get('m1');
    expect(msg.body?.content).toBe('old mail');
  });

  it('search: drops fresh non-draft hits; keeps drafts and older hits', async () => {
    const { graph } = fakeGraph([
      {
        value: [
          { id: 'fresh', isDraft: false, receivedDateTime: FRESH_RECEIVED, subject: 'OTP: 123456' },
          { id: 'draft', isDraft: true, receivedDateTime: FRESH_RECEIVED, subject: 'My draft' },
          { id: 'old', isDraft: false, receivedDateTime: OLD_RECEIVED, subject: 'Old mail' },
        ],
      },
    ]);
    const out = new OutlookClient(graph);
    const page = await out.mail.search({ query: 'anything' });
    expect(page.count).toBe(2);
    expect(page.results.map((r) => r.id)).toEqual(['draft', 'old']);
  });

  it('reply: pre-flight refuses fresh non-draft; no draft is created', async () => {
    const { graph, calls } = fakeGraph([
      { id: 'fresh', isDraft: false, receivedDateTime: FRESH_RECEIVED },
    ]);
    const out = new OutlookClient(graph);
    await expect(out.mail.reply('fresh', { body: 'hi' })).rejects.toThrow(
      /safety window/i,
    );
    // Only the pre-flight GET fired — no POST to createReply.
    expect(calls.capturedMethods).toEqual(['GET']);
  });

  it('forward: pre-flight refuses fresh non-draft; no draft is created', async () => {
    const { graph, calls } = fakeGraph([
      { id: 'fresh', isDraft: false, receivedDateTime: FRESH_RECEIVED },
    ]);
    const out = new OutlookClient(graph);
    await expect(
      out.mail.forward('fresh', { to: ['x@y.com'] }),
    ).rejects.toThrow(/safety window/i);
    expect(calls.capturedMethods).toEqual(['GET']);
  });

  it('downloadAttachment: pre-flight refuses fresh non-draft; bytes are never fetched', async () => {
    const { graph, calls } = fakeGraph([
      { id: 'fresh', isDraft: false, receivedDateTime: FRESH_RECEIVED },
    ]);
    const out = new OutlookClient(graph);
    await expect(
      out.mail.downloadAttachment('fresh', 'aid-1'),
    ).rejects.toThrow(/safety window/i);
    // Only the pre-flight ran; we never hit the attachment endpoint.
    expect(calls.capturedPaths).toEqual(['/me/messages/fresh']);
  });

  it('downloadAttachment: pre-flight passes for drafts (user attaching their own file)', async () => {
    const { graph } = fakeGraph([
      { id: 'd', isDraft: true, receivedDateTime: FRESH_RECEIVED },
      {
        '@odata.type': '#microsoft.graph.fileAttachment',
        id: 'a',
        name: 'doc.pdf',
        contentType: 'application/pdf',
        contentBytes: Buffer.from('x').toString('base64'),
      },
    ]);
    const out = new OutlookClient(graph);
    const r = await out.mail.downloadAttachment('d', 'a');
    expect(r.name).toBe('doc.pdf');
  });
});

describe('composeLinkFromWebLink', () => {
  it('extracts ItemID and URL-encodes it', () => {
    const web =
      'https://outlook.office.com/mail/inbox/id/AAA%2BBBB?ItemID=AQMkA%2FBcd%3D';
    expect(composeLinkFromWebLink(web)).toBe(
      'https://outlook.cloud.microsoft/mail/compose/AQMkA%2FBcd%3D',
    );
  });
  it('returns null when ItemID is absent', () => {
    expect(composeLinkFromWebLink('https://outlook.office.com/mail/id/123')).toBeNull();
  });
  it('returns null when webLink is empty / malformed', () => {
    expect(composeLinkFromWebLink(null)).toBeNull();
    expect(composeLinkFromWebLink(undefined)).toBeNull();
    expect(composeLinkFromWebLink('not a url')).toBeNull();
  });
});

describe('MailResource.draft', () => {
  it('POSTs /me/messages with text body and recipient envelope', async () => {
    const { graph, calls } = fakeGraph([
      {
        id: 'draft-1',
        subject: 'Hello',
        webLink: 'https://outlook.office.com/mail/drafts/id/x?ItemID=AAA%3D',
        toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
        ccRecipients: [{ emailAddress: { address: 'c@example.com' } }],
        bccRecipients: [],
      },
    ]);
    const out = new OutlookClient(graph);
    const summary = await out.mail.draft({
      subject: 'Hello',
      body: 'Hi there',
      to: ['a@example.com'],
      cc: ['c@example.com'],
    });
    expect(calls.capturedPaths).toEqual(['/me/messages']);
    expect(calls.capturedMethods).toEqual(['POST']);
    expect(calls.capturedPosts[0]).toMatchObject({
      subject: 'Hello',
      body: { contentType: 'Text', content: 'Hi there' },
      toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
      ccRecipients: [{ emailAddress: { address: 'c@example.com' } }],
      bccRecipients: [],
    });
    expect(summary.id).toBe('draft-1');
    expect(summary.to).toEqual(['a@example.com']);
    expect(summary.cc).toEqual(['c@example.com']);
    expect(summary.composeLink).toBe(
      'https://outlook.cloud.microsoft/mail/compose/AAA%3D',
    );
  });

  it('uses HTML content-type when html=true', async () => {
    const { graph, calls } = fakeGraph([
      { id: 'd2', subject: 's', toRecipients: [], ccRecipients: [], bccRecipients: [] },
    ]);
    const out = new OutlookClient(graph);
    await out.mail.draft({
      subject: 's',
      body: '<p>hi</p>',
      html: true,
      to: ['a@example.com'],
    });
    expect(calls.capturedPosts[0]).toMatchObject({
      body: { contentType: 'HTML', content: '<p>hi</p>' },
    });
  });

  it('rejects empty recipient list', async () => {
    const { graph } = fakeGraph([]);
    const out = new OutlookClient(graph);
    await expect(
      out.mail.draft({ subject: 's', body: 'b', to: [] }),
    ).rejects.toThrow('--to recipient');
  });
});

describe('MailResource.reply', () => {
  it('POSTs createReply then PATCHes the body', async () => {
    const seeded = {
      id: 'draft-reply',
      subject: 'RE: hi',
      webLink: 'https://outlook.office.com/mail/drafts/x?ItemID=BBB',
      toRecipients: [{ emailAddress: { address: 'sender@example.com' } }],
      ccRecipients: [],
      bccRecipients: [],
    };
    const { graph, calls } = fakeGraph([
      PRE_FLIGHT_OLD_INBOUND,
      seeded,
      { id: 'draft-reply', subject: 'RE: hi' },
    ]);
    const out = new OutlookClient(graph);
    const summary = await out.mail.reply('msg-1', { body: 'Got it.', html: false });
    expect(calls.capturedPaths).toEqual([
      '/me/messages/msg-1',
      '/me/messages/msg-1/createReply',
      '/me/messages/draft-reply',
    ]);
    expect(calls.capturedMethods).toEqual(['GET', 'POST', 'PATCH']);
    expect(calls.capturedPosts[0]).toEqual({});
    expect(calls.capturedPatches[0]).toMatchObject({
      body: { contentType: 'Text', content: 'Got it.' },
    });
    expect(summary.id).toBe('draft-reply');
    expect(summary.to).toEqual(['sender@example.com']);
    expect(summary.composeLink).toBe(
      'https://outlook.cloud.microsoft/mail/compose/BBB',
    );
  });

  it('routes to createReplyAll when replyAll=true', async () => {
    const { graph, calls } = fakeGraph([
      PRE_FLIGHT_OLD_INBOUND,
      { id: 'd', subject: 's', toRecipients: [], ccRecipients: [], bccRecipients: [] },
      { id: 'd' },
    ]);
    const out = new OutlookClient(graph);
    await out.mail.reply('m', { body: 'b', replyAll: true });
    expect(calls.capturedPaths[1]).toBe('/me/messages/m/createReplyAll');
  });
});

describe('MailResource.forward', () => {
  it('POSTs createForward with comment + recipients', async () => {
    const { graph, calls } = fakeGraph([
      PRE_FLIGHT_OLD_INBOUND,
      {
        id: 'fwd-1',
        subject: 'FW: hi',
        webLink: 'https://outlook.office.com/mail/drafts/x?ItemID=CCC',
        toRecipients: [{ emailAddress: { address: 'next@example.com' } }],
        ccRecipients: [{ emailAddress: { address: 'cc@example.com' } }],
        bccRecipients: [],
      },
    ]);
    const out = new OutlookClient(graph);
    const summary = await out.mail.forward('msg-1', {
      to: ['next@example.com'],
      cc: ['cc@example.com'],
      comment: 'fyi',
    });
    expect(calls.capturedPaths).toEqual([
      '/me/messages/msg-1',
      '/me/messages/msg-1/createForward',
    ]);
    expect(calls.capturedMethods).toEqual(['GET', 'POST']);
    expect(calls.capturedPosts[0]).toMatchObject({
      comment: 'fyi',
      toRecipients: [{ emailAddress: { address: 'next@example.com' } }],
      ccRecipients: [{ emailAddress: { address: 'cc@example.com' } }],
    });
    expect(summary.id).toBe('fwd-1');
    expect(summary.cc).toEqual(['cc@example.com']);
  });

  it('defaults comment to empty string', async () => {
    const { graph, calls } = fakeGraph([
      PRE_FLIGHT_OLD_INBOUND,
      { id: 'f', subject: 'FW', toRecipients: [], ccRecipients: [], bccRecipients: [] },
    ]);
    const out = new OutlookClient(graph);
    await out.mail.forward('m', { to: ['x@example.com'] });
    expect(calls.capturedPosts[0]).toMatchObject({ comment: '' });
  });

  it('rejects empty recipient list', async () => {
    const { graph } = fakeGraph([]);
    const out = new OutlookClient(graph);
    await expect(out.mail.forward('m', { to: [] })).rejects.toThrow('--to recipient');
  });
});

describe('MailResource.addAttachment', () => {
  it('POSTs base64-encoded fileAttachment payload', async () => {
    const { graph, calls } = fakeGraph([
      {
        id: 'att-1',
        name: 'hi.txt',
        contentType: 'text/plain',
        size: 5,
        isInline: false,
      },
    ]);
    const out = new OutlookClient(graph);
    const summary = await out.mail.addAttachment('draft-1', {
      name: 'hi.txt',
      contentType: 'text/plain',
      contentBytes: Buffer.from('hello'),
    });
    expect(calls.capturedPaths).toEqual(['/me/messages/draft-1/attachments']);
    expect(calls.capturedMethods).toEqual(['POST']);
    expect(calls.capturedPosts[0]).toMatchObject({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'hi.txt',
      contentType: 'text/plain',
      contentBytes: Buffer.from('hello').toString('base64'),
      isInline: false,
    });
    expect(summary.attachmentId).toBe('att-1');
    expect(summary.size).toBe(5);
  });

  it('honours isInline=true', async () => {
    const { graph, calls } = fakeGraph([
      { id: 'a', name: 'img.png', contentType: 'image/png', size: 1, isInline: true },
    ]);
    const out = new OutlookClient(graph);
    await out.mail.addAttachment('d', {
      name: 'img.png',
      contentType: 'image/png',
      contentBytes: Buffer.from([0]),
      isInline: true,
    });
    expect(calls.capturedPosts[0]).toMatchObject({ isInline: true });
  });

  it('rejects files larger than 3 MB', async () => {
    const { graph } = fakeGraph([]);
    const out = new OutlookClient(graph);
    const big = Buffer.alloc(3 * 1024 * 1024 + 1);
    await expect(
      out.mail.addAttachment('d', { name: 'big.bin', contentType: 'application/octet-stream', contentBytes: big }),
    ).rejects.toThrow('createUploadSession');
  });
});

describe('MailResource.move', () => {
  it('POSTs /me/messages/<id>/move with well-known folder destination', async () => {
    const { graph, calls } = fakeGraph([
      { id: 'new-msg-id', subject: 'hi' },
    ]);
    const out = new OutlookClient(graph);
    const result = await out.mail.move('old-id', 'archive');
    expect(calls.capturedPaths).toEqual(['/me/messages/old-id/move']);
    expect(calls.capturedMethods).toEqual(['POST']);
    expect(calls.capturedPosts[0]).toEqual({ destinationId: 'archive' });
    expect(result).toEqual({
      id: 'new-msg-id',
      oldId: 'old-id',
      destinationFolder: 'archive',
    });
  });

  it('resolves a custom folder displayName before moving', async () => {
    const { graph, calls } = fakeGraph([
      { value: [{ id: 'fld-xyz', displayName: 'Projects' }] },
      { id: 'new-msg-id', subject: 'p' },
    ]);
    const out = new OutlookClient(graph);
    const result = await out.mail.move('msg-1', 'Projects');
    expect(calls.capturedPaths).toEqual([
      '/me/mailFolders',
      '/me/messages/msg-1/move',
    ]);
    expect(calls.capturedPosts[0]).toEqual({ destinationId: 'fld-xyz' });
    expect(result.destinationFolder).toBe('fld-xyz');
    expect(result.id).toBe('new-msg-id');
  });
});

describe('MailResource.delete', () => {
  it('DELETEs /me/messages/<id> and returns deleted: true', async () => {
    const { graph, calls } = fakeGraph([undefined]);
    const out = new OutlookClient(graph);
    const result = await out.mail.delete('msg-1');
    expect(calls.capturedPaths).toEqual(['/me/messages/msg-1']);
    expect(calls.capturedMethods).toEqual(['DELETE']);
    expect(result).toEqual({ id: 'msg-1', deleted: true });
  });
});

describe('MailResource.mark', () => {
  it('PATCHes isRead: true when read', async () => {
    const { graph, calls } = fakeGraph([undefined]);
    const out = new OutlookClient(graph);
    const result = await out.mail.mark('msg-1', true);
    expect(calls.capturedPaths).toEqual(['/me/messages/msg-1']);
    expect(calls.capturedMethods).toEqual(['PATCH']);
    expect(calls.capturedPatches[0]).toEqual({ isRead: true });
    expect(result).toEqual({ id: 'msg-1', isRead: true });
  });

  it('PATCHes isRead: false when unread', async () => {
    const { graph, calls } = fakeGraph([undefined]);
    const out = new OutlookClient(graph);
    const result = await out.mail.mark('msg-1', false);
    expect(calls.capturedPatches[0]).toEqual({ isRead: false });
    expect(result.isRead).toBe(false);
  });
});

describe('MailResource.flag', () => {
  it('PATCHes flag.flagStatus', async () => {
    const { graph, calls } = fakeGraph([undefined]);
    const out = new OutlookClient(graph);
    const result = await out.mail.flag('msg-1', 'flagged');
    expect(calls.capturedPaths).toEqual(['/me/messages/msg-1']);
    expect(calls.capturedMethods).toEqual(['PATCH']);
    expect(calls.capturedPatches[0]).toEqual({ flag: { flagStatus: 'flagged' } });
    expect(result).toEqual({ id: 'msg-1', flagStatus: 'flagged' });
  });
});

describe('MailResource.importance', () => {
  it('PATCHes importance level', async () => {
    const { graph, calls } = fakeGraph([undefined]);
    const out = new OutlookClient(graph);
    const result = await out.mail.importance('msg-1', 'high');
    expect(calls.capturedPaths).toEqual(['/me/messages/msg-1']);
    expect(calls.capturedMethods).toEqual(['PATCH']);
    expect(calls.capturedPatches[0]).toEqual({ importance: 'high' });
    expect(result).toEqual({ id: 'msg-1', importance: 'high' });
  });
});
