import type { AccountInfo } from '@azure/msal-node';

/** Result of a successful interactive / authorization-code sign-in. */
export interface LoginResult {
  account: AccountInfo;
  expiresOn: Date;
}
