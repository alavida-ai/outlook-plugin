import type { Client } from '@microsoft/microsoft-graph-client';

import { CalendarResource } from './resources/calendar.js';
import { MailResource } from './resources/mail.js';
import { MeResource } from './resources/me.js';

/**
 * Top-level facade over the Graph endpoints we expose. New resources
 * (contacts) hang off this client in subsequent slices.
 */
export class OutlookClient {
  public readonly me: MeResource;
  public readonly mail: MailResource;
  public readonly calendar: CalendarResource;

  constructor(graph: Client) {
    this.me = new MeResource(graph);
    this.mail = new MailResource(graph);
    this.calendar = new CalendarResource(graph);
  }
}
