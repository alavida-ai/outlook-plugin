/**
 * OpenClaw plugin entry — outlook.
 *
 * Reads and triages Outlook mail + calendar via Microsoft Graph as the
 * signed-in user. Delegated permissions; draft-only mail.
 *
 * Architecture:
 *   - One file per tool in `./tools/<tool-name>.ts`, each default-exporting
 *     a {@link ToolDescriptor}.
 *   - {@link registerTool} wraps every descriptor with shared output/help
 *     injection, pretty/json dispatch, and `withErrorMapping`.
 *
 * This first slice ships `whoami` only. Mail, calendar, and contacts tools
 * land in subsequent plan files.
 */
import { Type } from 'typebox';
import { definePluginEntry } from 'openclaw/plugin-sdk/plugin-entry';

import type { PluginConfig } from './client.js';
import { readPluginConfig, registerTool, type ToolDescriptor } from './register.js';

import authLogin from './tools/auth-login.js';
import authStatus from './tools/auth-status.js';
import authLogout from './tools/auth-logout.js';
import whoami from './tools/whoami.js';
import mailList from './tools/mail-list.js';
import mailRead from './tools/mail-read.js';
import mailSearch from './tools/mail-search.js';
import mailFolders from './tools/mail-folders.js';
import mailListAttachments from './tools/mail-list-attachments.js';
import mailDownloadAttachment from './tools/mail-download-attachment.js';
import mailDraft from './tools/mail-draft.js';
import mailReply from './tools/mail-reply.js';
import mailForward from './tools/mail-forward.js';
import mailAddAttachment from './tools/mail-add-attachment.js';
import calendarList from './tools/calendar-list.js';
import calendarShow from './tools/calendar-show.js';
import calendarAvailability from './tools/calendar-availability.js';

/**
 * Tools deliberately not registered (see
 * clients/sunglobal/scoping/agent-data-handling-response.md):
 *   - mail send/move/delete/mark/flag/importance — Mail.ReadWrite scope
 *     would permit them but the agent surface is draft-only by design.
 *   - calendar create/update/delete/respond — read-only scope; writes
 *     would 403 at Graph even if registered.
 *   - contacts — `Contacts.ReadWrite` is not in the requested scope set.
 */
const TOOLS: ToolDescriptor[] = [
  authLogin,
  authStatus,
  authLogout,
  whoami,
  mailList,
  mailRead,
  mailSearch,
  mailFolders,
  mailListAttachments,
  mailDownloadAttachment,
  mailDraft,
  mailReply,
  mailForward,
  mailAddAttachment,
  calendarList,
  calendarShow,
  calendarAvailability,
];

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const configJsonSchema: any = Type.Object({
  clientId: Type.Optional(
    Type.String({ description: 'Override the embedded Entra app id.' }),
  ),
  tenantId: Type.Optional(
    Type.String({ description: "Override the default tenant ('common')." }),
  ),
  tokenCachePath: Type.Optional(
    Type.String({
      description:
        'Path to the MSAL token cache. Defaults to <agentDir>/outlook-tokens.json on a multi-agent OpenClaw gateway (per-agent isolation), or ~/.outlook-plugin/tokens.json when no agent context is available. Override only when you need explicit control.',
    }),
  ),
  account: Type.Optional(
    Type.String({
      description: 'UPN to use when multiple accounts are cached. Required for multi-account hosts.',
    }),
  ),
});

export default definePluginEntry({
  id: 'outlook',
  name: 'outlook',
  description:
    'Read and triage Outlook mail + calendar via Microsoft Graph as the signed-in user. Delegated permissions; draft-only mail.',
  configSchema: { jsonSchema: configJsonSchema },
  register(api) {
    const getConfig = () => readPluginConfig(api);
    for (const tool of TOOLS) {
      registerTool(api, tool, getConfig);
    }
  },
});

export { getClient, _resetClientForTesting } from './client.js';
export type { PluginConfig };
export { withErrorMapping, isToolErrorEnvelope } from './errors.js';
export type { ToolErrorEnvelope, ToolErrorResponse } from './errors.js';
export { registerTool, defineTool, type ToolDescriptor } from './register.js';
