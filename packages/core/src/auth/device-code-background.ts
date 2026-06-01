import type { DeviceCodeRequest, PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import { OUTLOOK_SCOPES } from './msal.js';

// @azure/msal-node does not re-export DeviceCodeResponse from msal-common; we
// derive it from the callback signature on DeviceCodeRequest so we don't have
// to take a direct dep on @azure/msal-common just for one type.
type DeviceCodeResponse = Parameters<DeviceCodeRequest['deviceCodeCallback']>[0];

export interface LoginDeviceCodeBackgroundInput {
  app: PublicClientApplication;
  cache: TokenCache;
}

export interface BackgroundDeviceCodeResult {
  /** Microsoft device-login URL the user opens in a browser. */
  verificationUrl: string;
  /** Six-character code the user enters at the URL. */
  userCode: string;
  /** ISO timestamp when the device code expires (Microsoft default ~15 min). */
  expiresAt: string;
  /**
   * Resolves when MSAL polling completes. Plugin callers don't need to await
   * this — `auth_status` is the user-facing confirmation. Exposed for tests
   * and for callers who want to log silent failures.
   */
  completion: Promise<BackgroundDeviceCodeCompletion>;
}

export type BackgroundDeviceCodeCompletion =
  | { ok: true; upn: string }
  | { ok: false; reason: string };

/**
 * Non-blocking variant of {@link loginDeviceCode}. Kicks off MSAL device-code
 * acquisition, resolves as soon as Microsoft emits the user-facing
 * `verificationUrl` + `userCode` (~200 ms), and continues polling in the
 * background until the user signs in or MSAL times out.
 *
 * Used by the OpenClaw plugin's `auth_login` tool: the agent surfaces the URL
 * and code to the human immediately, then later calls `auth_status` to
 * confirm. The CLI keeps using the blocking `loginDeviceCode` because the
 * terminal naturally pauses anyway.
 */
export function loginDeviceCodeInBackground(
  input: LoginDeviceCodeBackgroundInput,
): Promise<BackgroundDeviceCodeResult> {
  const { app, cache } = input;

  return new Promise<BackgroundDeviceCodeResult>((resolveDeviceCode, rejectDeviceCode) => {
    // Latch state for the device-code response: once MSAL invokes the
    // callback, resolve the outer promise. If MSAL throws before then,
    // reject — we have nothing to hand back to the caller.
    let deviceCodeSeen = false;

    const completion: Promise<BackgroundDeviceCodeCompletion> = cache
      .lock(async () =>
        app.acquireTokenByDeviceCode({
          scopes: [...OUTLOOK_SCOPES],
          deviceCodeCallback: (info: DeviceCodeResponse) => {
            deviceCodeSeen = true;
            const expiresAt = new Date(Date.now() + info.expiresIn * 1000).toISOString();
            resolveDeviceCode({
              verificationUrl: info.verificationUri,
              userCode: info.userCode,
              expiresAt,
              completion,
            });
          },
        }),
      )
      .then((result): BackgroundDeviceCodeCompletion => {
        if (!result?.account) {
          return { ok: false, reason: 'device-code flow returned no account' };
        }
        return { ok: true, upn: result.account.username };
      })
      .catch((err): BackgroundDeviceCodeCompletion => {
        // If MSAL never even emitted the device code (e.g. network failure),
        // the outer promise hasn't resolved yet — reject it so callers get a
        // synchronous failure rather than a stuck "pending" envelope.
        if (!deviceCodeSeen) rejectDeviceCode(err);
        const message = err instanceof Error ? err.message : String(err);
        return { ok: false, reason: message };
      });

    // Don't let the completion promise's rejection surface as an
    // unhandled-rejection in callers that intentionally ignore it.
    completion.catch(() => {});
  });
}
