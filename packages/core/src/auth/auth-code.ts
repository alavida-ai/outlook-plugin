import {
  CryptoProvider,
  type AuthorizationCodeRequest,
  type AuthorizationUrlRequest,
  type PublicClientApplication,
} from '@azure/msal-node';

import type { TokenCache } from './cache.js';
import type { LoginResult } from './login-result.js';
import { AuthRefreshFailedError } from './errors.js';
import { OUTLOOK_SCOPES } from './msal.js';

/**
 * How long a pending Authorization-Code flow stays valid between
 * {@link buildAuthCodeUrl} (URL handed to the user) and the browser callback
 * that calls {@link exchangeAuthCode}. After this the server-side flow entry is
 * garbage-collected and the callback is refused. Microsoft's own auth-code
 * lifetime is ~10 min, so matching it here keeps the two windows aligned.
 */
export const PENDING_FLOW_TTL_MS = 10 * 60 * 1000;

export interface BuildAuthCodeUrlInput {
  app: PublicClientApplication;
  /**
   * Public HTTPS redirect URI registered in the Entra app. Microsoft requires
   * an exact match against the app registration.
   */
  redirectUri: string;
  /** Override the default delegated scope set (mainly for tests). */
  scopes?: readonly string[];
}

export interface AuthCodeUrlResult {
  /** The login.microsoftonline.com URL the user opens in a browser. */
  authUrl: string;
  /**
   * CSRF token. Caller stores it server-side keyed to the initiating agent and
   * matches it against the `state` query param on the callback. Never enters
   * LLM context as a secret — the agent only forwards the opaque `authUrl`.
   */
  state: string;
  /**
   * PKCE code verifier. MUST stay server-side; only its SHA-256 challenge is
   * sent to Microsoft. Caller passes it back into {@link exchangeAuthCode}.
   */
  verifier: string;
  /** ID-token replay guard. Microsoft binds it into the issued ID token. */
  nonce: string;
  /** ISO timestamp after which the pending flow must be refused. */
  expiresAt: string;
}

/**
 * First leg of the OAuth 2.0 Authorization Code flow with PKCE.
 *
 * Generates a PKCE verifier/challenge pair, a CSRF `state`, and an ID-token
 * `nonce`, then asks MSAL to build the authorize URL. The caller surfaces
 * `authUrl` to the user and stashes `{ state, verifier, nonce, expiresAt }`
 * server-side until the browser redirect arrives.
 *
 * This does no network round-trip of its own — MSAL only assembles a URL —
 * so there is nothing to lock here.
 */
export async function buildAuthCodeUrl(
  input: BuildAuthCodeUrlInput,
): Promise<AuthCodeUrlResult> {
  const { app, redirectUri } = input;
  const scopes = [...(input.scopes ?? OUTLOOK_SCOPES)];

  const crypto = new CryptoProvider();
  const { verifier, challenge } = await crypto.generatePkceCodes();
  const state = crypto.createNewGuid();
  const nonce = crypto.createNewGuid();

  const request: AuthorizationUrlRequest = {
    scopes,
    redirectUri,
    codeChallenge: challenge,
    codeChallengeMethod: 'S256',
    state,
    nonce,
  };
  const authUrl = await app.getAuthCodeUrl(request);

  return {
    authUrl,
    state,
    verifier,
    nonce,
    expiresAt: new Date(Date.now() + PENDING_FLOW_TTL_MS).toISOString(),
  };
}

export interface ExchangeAuthCodeInput {
  app: PublicClientApplication;
  cache: TokenCache;
  /** The `code` query param Microsoft redirected back with. */
  code: string;
  /** The PKCE verifier minted in {@link buildAuthCodeUrl} for this flow. */
  verifier: string;
  /** Must be byte-for-byte the redirect URI used to build the auth URL. */
  redirectUri: string;
  /** The `nonce` minted in {@link buildAuthCodeUrl}; validated against the ID token. */
  expectedNonce: string;
  /** The `state` minted in {@link buildAuthCodeUrl}, echoed to MSAL for completeness. */
  expectedState?: string;
  scopes?: readonly string[];
}

/**
 * Second leg of the Authorization Code flow: redeem `code` + PKCE `verifier`
 * for tokens, writing them through the cache plugin (so per-agent isolation is
 * honored by whichever {@link TokenCache} the caller built the app with).
 *
 * Holds the cache lock for the full read-modify-write, matching the rest of the
 * auth layer (spec §4.2.2). Verifies the ID-token nonce echoes the one we
 * issued — a mismatch means the ID token isn't the one minted for this flow.
 */
export async function exchangeAuthCode(input: ExchangeAuthCodeInput): Promise<LoginResult> {
  const { app, cache, code, verifier, redirectUri, expectedNonce, expectedState } = input;
  const scopes = [...(input.scopes ?? OUTLOOK_SCOPES)];

  return cache.lock(async () => {
    const request: AuthorizationCodeRequest = {
      scopes,
      redirectUri,
      code,
      codeVerifier: verifier,
      state: expectedState,
    };
    const result = await app.acquireTokenByCode(request);
    if (!result?.account || !result.expiresOn) {
      throw new AuthRefreshFailedError('auth-code flow returned no account');
    }

    const nonce = (result.idTokenClaims as { nonce?: string } | undefined)?.nonce;
    if (nonce !== expectedNonce) {
      throw new AuthRefreshFailedError('ID token nonce did not match the issued nonce');
    }

    return { account: result.account, expiresOn: result.expiresOn };
  });
}
