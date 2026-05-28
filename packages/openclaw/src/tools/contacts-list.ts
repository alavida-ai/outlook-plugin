/**
 * `contacts_list` — stub. Mirrors the Python implementation, which has not
 * been wired to Graph /me/contacts yet. Returns a static stub envelope so
 * agents see a clear "not yet implemented" rather than an error.
 */
import { Type } from 'typebox';

import { defineTool } from '../register.js';

interface ContactsListStub {
  stub: true;
  message: string;
}

const contactsList = defineTool({
  name: 'contacts_list',
  description:
    'STUB — not yet implemented. Real Graph /me/contacts port lives behind a future ticket. Returns a sentinel envelope so agents can detect and route around it.',
  parameters: Type.Object({}),
  async execute(_params, _config): Promise<ContactsListStub> {
    return {
      stub: true,
      message: 'contacts_list is a stub; not yet implemented.',
    };
  },
});

export default contactsList;
