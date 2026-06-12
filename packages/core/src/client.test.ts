import { describe, expect, it, vi } from 'vitest';
import type { Client } from '@microsoft/microsoft-graph-client';

import { OutlookClient } from './client.js';

function fakeGraphClient(meResponse: unknown): Client {
  return {
    api: vi.fn(() => ({
      get: vi.fn(async () => meResponse),
    })),
  } as unknown as Client;
}

describe('OutlookClient.me.get', () => {
  it('returns the Graph /me payload as-is', async () => {
    const sample = {
      id: 'user-id',
      displayName: 'Alice Example',
      mail: 'alice@example.com',
      userPrincipalName: 'alice@example.com',
      jobTitle: 'Engineer',
      department: 'Eng',
      officeLocation: 'London',
    };
    const client = new OutlookClient(fakeGraphClient(sample));
    expect(await client.me.get()).toEqual(sample);
  });

  it('passes /me to .api()', async () => {
    const graph = fakeGraphClient({});
    const client = new OutlookClient(graph);
    await client.me.get();
    expect(graph.api).toHaveBeenCalledWith('/me');
  });
});
