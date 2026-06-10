/**
 * `whoami` — display the signed-in user's profile.
 *
 * Mirrors `outlook whoami` on the CLI. Pulls /me from Graph and returns a
 * compact subset (id, displayName, mail, userPrincipalName, jobTitle,
 * department, officeLocation). Read-only.
 */
import { Type } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

interface WhoamiResult {
  id: string | null;
  displayName: string | null;
  mail: string | null;
  userPrincipalName: string | null;
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
}

const whoami = defineTool({
  name: 'outlook_whoami',
  description:
    "Display the signed-in user's basic profile (display name, email, job title, department, office). Read-only.",
  parameters: Type.Object({}),
  async execute(_params, config): Promise<WhoamiResult> {
    const client = getClient(config);
    const me = await client.me.get();
    return {
      id: me.id ?? null,
      displayName: me.displayName ?? null,
      mail: me.mail ?? null,
      userPrincipalName: me.userPrincipalName ?? null,
      jobTitle: me.jobTitle ?? null,
      department: me.department ?? null,
      officeLocation: me.officeLocation ?? null,
    };
  },
});

export default whoami;
