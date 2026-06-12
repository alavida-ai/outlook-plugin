import { spawn } from 'node:child_process';
import { parseArgs } from 'node:util';
import { loginInteractive } from '@alavida-ai/outlook-core';

import { makeContext } from '../client.js';
import { eprintln, formatError, printJson, println } from '../output.js';

const HELP = `Usage: outlook auth login [--json]

Sign in via the browser (Authorization Code + PKCE) and cache the tokens.

  - Your default browser opens to the Microsoft sign-in page. If it can't be
    opened (e.g. over SSH), the URL is printed to STDERR — open it yourself.
  - MSAL listens on a localhost loopback for the redirect; nothing needs to be
    exposed publicly and no redirect URI has to be registered.
  - Tokens are cached at ~/.outlook-plugin/tokens.json (0600) — set
    OUTLOOK_TOKEN_CACHE to override.
  - With --json, the result is emitted to stdout as
    { "account": "<upn>" }.
`;

/**
 * Open `url` in the system browser. Best-effort: we also print the URL to
 * STDERR so a headless user can copy it. Spawn failures are swallowed — the
 * printed URL is the fallback.
 */
async function openInBrowser(url: string): Promise<void> {
  eprintln(`Opening your browser to sign in. If it doesn't open, visit:\n${url}`);
  const command =
    process.platform === 'darwin'
      ? 'open'
      : process.platform === 'win32'
        ? 'cmd'
        : 'xdg-open';
  const args = process.platform === 'win32' ? ['/c', 'start', '', url] : [url];
  try {
    const child = spawn(command, args, { stdio: 'ignore', detached: true });
    child.on('error', () => {});
    child.unref();
  } catch {
    // Ignore — the URL was already printed for manual use.
  }
}

export async function run(argv: string[]): Promise<number> {
  let parsed;
  try {
    parsed = parseArgs({
      args: argv,
      options: {
        json: { type: 'boolean', default: false },
        help: { type: 'boolean', default: false, short: 'h' },
      },
      strict: true,
    });
  } catch (err) {
    eprintln(formatError(err));
    eprintln(HELP);
    return 1;
  }

  if (parsed.values.help) {
    process.stdout.write(HELP);
    return 0;
  }

  const ctx = makeContext();

  try {
    const result = await loginInteractive({
      app: ctx.app,
      cache: ctx.cache,
      openBrowser: openInBrowser,
      successTemplate:
        '<h1>Signed in</h1><p>You can close this tab and return to your terminal.</p>',
      errorTemplate:
        '<h1>Sign-in failed</h1><p>Return to your terminal and try again.</p>',
    });
    if (parsed.values.json) {
      printJson({ account: result.account.username });
    } else {
      println(`Signed in as ${result.account.username}.`);
    }
    return 0;
  } catch (err) {
    eprintln(formatError(err));
    return 1;
  }
}
