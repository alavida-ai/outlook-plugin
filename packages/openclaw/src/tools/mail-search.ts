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
  name: 'outlook_mail_search',
  description:
    'Search Outlook messages across all folders using a KQL query. Read-only. Results are relevance-ranked (not chronological). Inbound non-draft hits less than 30 min old are filtered out (safety window for one-time passwords) — count reflects the post-filter view. Each hit carries an inboxLink (outlook.cloud.microsoft URL) for surfacing to the user.',
  parameters: Params,
  async execute(params: Static<typeof Params>, config) {
    const client = getClient(config);
    const page = await client.mail.search({ query: params.query, limit: params.limit });
    // Batch-translate REST ids to restImmutableEntryId for cloud URLs.
    const restIds = page.results
      .map((m) => m.id)
      .filter((id): id is string => typeof id === 'string' && id.length > 0);
    let inboxLinks: Record<string, string | null> = {};
    if (restIds.length > 0) {
      try {
        inboxLinks = await client.mail.inboxLinks(restIds);
      } catch {
        // Translation failures are non-fatal.
      }
    }
    return {
      messages: page.results.map((m) => ({
        id: m.id ?? null,
        subject: m.subject ?? null,
        from: m.from?.emailAddress?.address ?? null,
        receivedDateTime: m.receivedDateTime ?? null,
        isRead: m.isRead ?? null,
        isDraft: m.isDraft ?? null,
        hasAttachments: m.hasAttachments ?? null,
        bodyPreview: m.bodyPreview ?? null,
        webLink: m.webLink ?? null,
        inboxLink: m.id ? (inboxLinks[m.id] ?? null) : null,
      })),
      count: page.count,
      nextLink: page.nextLink,
    };
  },
});

export default mailSearch;
