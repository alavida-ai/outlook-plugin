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

import whoami from './tools/whoami.js';
import mailList from './tools/mail-list.js';
import mailRead from './tools/mail-read.js';
import mailSearch from './tools/mail-search.js';
import mailFolders from './tools/mail-folders.js';
import mailListAttachments from './tools/mail-list-attachments.js';
import mailDownloadAttachment from './tools/mail-download-attachment.js';

const TOOLS: ToolDescriptor[] = [
  whoami,
  mailList,
  mailRead,
  mailSearch,
  mailFolders,
  mailListAttachments,
  mailDownloadAttachment,
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
      description: 'Path to the MSAL token cache. Defaults to ~/.outlook-cli/tokens.json.',
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
