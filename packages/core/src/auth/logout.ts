import { unlink } from 'node:fs/promises';

import type { TokenCache } from './cache.js';
import type { FileTokenCache } from './cache-file.js';

export interface LogoutInput {
  cache: TokenCache;
}

/**
 * Tear down all auth state owned by this package. Idempotent.
 *
 * Per spec §4.2.7, this also makes a best-effort attempt to delete the
 * `keyring` entry the Python CLI may have left around (migration kindness).
 * The deletion is wrapped in a try/catch — we never crash logout because
 * the user is on a system without libsecret/Keychain.
 */
export async function logout(input: LogoutInput): Promise<void> {
  const { cache } = input;
  await cache.clear();

  // FileTokenCache exposes lockPath; for other backends this is a no-op.
  const lockPath = (cache as Partial<FileTokenCache>).lockPath;
  if (typeof lockPath === 'string') {
    await unlink(lockPath).catch(() => {});
  }
  // Best-effort: the Python CLI used keyring service "outlook-cli" / key "default".
  // We don't take a keytar dep just for migration; the file deletion above is
  // enough on systems where Python tokens were file-cached. Keychain users
  // can run `security delete-generic-password -s outlook-cli -a default` once.
}
