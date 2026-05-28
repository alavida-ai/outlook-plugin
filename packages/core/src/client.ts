import type { Client } from '@microsoft/microsoft-graph-client';

import { MailResource } from './resources/mail.js';
import { MeResource } from './resources/me.js';

/**
 * Top-level facade over the Graph endpoints we expose. New resources
 * (calendar, contacts) hang off this client in subsequent slices.
 */
export class OutlookClient {
  public readonly me: MeResource;
  public readonly mail: MailResource;

  constructor(graph: Client) {
    this.me = new MeResource(graph);
    this.mail = new MailResource(graph);
  }
}
