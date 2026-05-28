/**
 * `mail_search` — KQL search across all folders.
 *
 * Sends `$search="<kql>"` (Graph requires the double quotes). Note that
 * `$search` and `$orderby` are mutually exclusive in Graph, so result order
 * is relevance-ranked, not chronological.
 */
import { Type, type Static } from 'typebox';

import { getClient } from '../client.js';
import { defineTool } from '../register.js';

const Params = Type.Object({
  query: Type.String({
    description:
      'KQL query (e.g. `from:boss@co.com subject:urgent`). Searches across every mail folder.',
  }),
  limit: Type.Optional(
    Type.Integer({
      minimum: 1,
      maximum: 200,
      default: 25,
      description: 'Maximum number of results to return (default 25).',
    }),
  ),
});

const mailSearch = defineTool({
  name: 'mail_search',
  description:
    'Search Outlook messages across all folders using a KQL query. Read-only. Results are relevance-ranked (not chronological).',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const page = await client.mail.search({ query: params.query, limit: params.limit });
    return {
      messages: page.results.map((m) => ({
        id: m.id ?? null,
        subject: m.subject ?? null,
        from: m.from?.emailAddress?.address ?? null,
        receivedDateTime: m.receivedDateTime ?? null,
        isRead: m.isRead ?? null,
        hasAttachments: m.hasAttachments ?? null,
        bodyPreview: m.bodyPreview ?? null,
        webLink: m.webLink ?? null,
      })),
      count: page.count,
      nextLink: page.nextLink,
    };
  },
});

export default mailSearch;
