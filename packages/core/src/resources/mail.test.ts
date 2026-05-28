import { describe, expect, it, vi } from 'vitest';
import { Client } from '@microsoft/microsoft-graph-client';

import { OutlookClient } from '../client.js';
import {
  normaliseDateForFilter,
  sanitiseAttachmentName,
} from './mail.js';

interface FakeRequest {
  api: ReturnType<typeof vi.fn>;
  query: ReturnType<typeof vi.fn>;
  header: ReturnType<typeof vi.fn>;
  get: ReturnType<typeof vi.fn>;
  capturedPaths: string[];
  capturedQueries: Array<Record<string, unknown>>;
  capturedHeaders: Array<[string, string]>;
}

/**
 * Build a fake Graph client that records the .api() path, all .query() args
 * and .header() args, and returns the supplied responses in order.
 *
 * Each call to .api() returns a chainable object whose .get() resolves to
 * the next queued response.
 */
function fakeGraph(responses: unknown[]): { graph: Client; calls: FakeRequest } {
  const queue = [...responses];
  const calls: FakeRequest = {
    api: vi.fn(),
    query: vi.fn(),
    header: vi.fn(),
    get: vi.fn(),
    capturedPaths: [],
    capturedQueries: [],
    capturedHeaders: [],
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
      calls.get();
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
    expect(q.$filter).toBe(
      "isRead eq false and from/emailAddress/address eq 'boss@example.com' and receivedDateTime ge 2026-05-01T00:00:00Z and receivedDateTime le 2026-05-28T00:00:00Z and inferenceClassification eq 'focused'",
    );
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
      {
        value: [
          { id: 'a1', name: 'file.pdf', contentType: 'application/pdf', size: 100, isInline: false },
        ],
      },
    ]);
    const out = new OutlookClient(graph);
    const page = await out.mail.listAttachments('msg-1');
    expect(calls.capturedPaths).toEqual(['/me/messages/msg-1/attachments']);
    expect(page.results[0]?.name).toBe('file.pdf');
  });
});

describe('MailResource.downloadAttachment', () => {
  it('decodes contentBytes and returns metadata', async () => {
    const payload = Buffer.from('hello world');
    const { graph, calls } = fakeGraph([
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
    expect(calls.capturedPaths).toEqual(['/me/messages/msg-1/attachments/a1']);
    expect(result.name).toBe('hello.txt');
    expect(result.contentType).toBe('text/plain');
    expect(result.size).toBe(payload.byteLength);
    expect(Buffer.compare(result.contentBytes, payload)).toBe(0);
  });
  it('rejects non-fileAttachment kinds', async () => {
    const { graph } = fakeGraph([
      { '@odata.type': '#microsoft.graph.itemAttachment', id: 'a2', name: 'card' },
    ]);
    const out = new OutlookClient(graph);
    await expect(out.mail.downloadAttachment('m', 'a2')).rejects.toThrow(
      'Cannot download attachment',
    );
  });
  it('sanitises the returned name', async () => {
    const { graph } = fakeGraph([
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
