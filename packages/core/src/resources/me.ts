import type { Client } from '@microsoft/microsoft-graph-client';
import type { User } from '@microsoft/microsoft-graph-types';

import { liftGraphError } from '../graph/errors.js';

/** Subset of `User` we surface today. The full Graph type is re-exported. */
export type MeProfile = Pick<
  User,
  | 'id'
  | 'displayName'
  | 'mail'
  | 'userPrincipalName'
  | 'jobTitle'
  | 'department'
  | 'officeLocation'
>;

export class MeResource {
  constructor(private readonly graph: Client) {}

  /** GET /me — basic profile for the signed-in user. */
  async get(): Promise<MeProfile> {
    try {
      return (await this.graph.api('/me').get()) as MeProfile;
    } catch (err) {
      throw liftGraphError(err);
    }
  }
}
