import type { AccountInfo, DeviceCodeRequest, PublicClientApplication } from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import { AuthRefreshFailedError } from './errors.js';
import { OUTLOOK_SCOPES } from './msal.js';

// @azure/msal-node does not re-export DeviceCodeResponse from msal-common; we
// derive it from the callback signature on DeviceCodeRequest so we don't have
// to take a direct dep on @azure/msal-common just for one type.
type DeviceCodeResponse = Parameters<DeviceCodeRequest['deviceCodeCallback']>[0];

export interface LoginDeviceCodeInput {
  app: PublicClientApplication;
  cache: TokenCache;
  /**
   * Called once when MSAL has the device-code response. The CLI prints
   * `info.message` to stderr verbatim — that text contains the URL and
   * the user code Microsoft expects you to display.
   */
  onDeviceCode: (info: DeviceCodeResponse) => void;
}

export interface LoginResult {
  account: AccountInfo;
  expiresOn: Date;
}

/**
 * Run the device-code flow under the cache lock. Blocks (up to MSAL's default
 * polling window — ~15 min) until the user finishes sign-in in their browser,
 * then writes the resulting tokens through the cache plugin.
 */
export async function loginDeviceCode(input: LoginDeviceCodeInput): Promise<LoginResult> {
  const { app, cache, onDeviceCode } = input;
  return cache.lock(async () => {
    const result = await app.acquireTokenByDeviceCode({
      scopes: [...OUTLOOK_SCOPES],
      deviceCodeCallback: (info) => onDeviceCode(info),
    });
    if (!result?.account || !result.expiresOn) {
      throw new AuthRefreshFailedError('device-code flow returned no account');
    }
    return { account: result.account, expiresOn: result.expiresOn };
  });
}
