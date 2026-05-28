import type { Client } from '@microsoft/microsoft-graph-client';

import { MeResource } from './resources/me.js';

/**
 * Top-level facade over the Graph endpoints we expose. New resources
 * (mail, calendar, contacts) hang off this client in subsequent slices.
 */
export class OutlookClient {
  public readonly me: MeResource;

  constructor(graph: Client) {
    this.me = new MeResource(graph);
  }
}
