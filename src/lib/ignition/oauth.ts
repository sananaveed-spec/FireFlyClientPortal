import { createHash, randomBytes } from "crypto";
import {
  assertIgnitionConfigured,
  getIgnitionConfig,
  type IgnitionConfig,
} from "@/lib/ignition/config";
import type { IgnitionTokenSet } from "@/lib/ignition/tokens";

export type OAuthTokenResponse = {
  access_token: string;
  refresh_token?: string;
  token_type?: string;
  expires_in?: number;
  scope?: string;
};

export type PkcePair = {
  codeVerifier: string;
  codeChallenge: string;
};

export function createPkcePair(): PkcePair {
  const codeVerifier = randomBytes(48).toString("base64url");
  const codeChallenge = createHash("sha256")
    .update(codeVerifier)
    .digest("base64url");
  return { codeVerifier, codeChallenge };
}

export function buildAuthorizeUrl(
  state: string,
  codeChallenge: string,
  config?: IgnitionConfig,
): string {
  const cfg = config ?? getIgnitionConfig();
  assertIgnitionConfigured(cfg);

  const url = new URL(cfg.authUrl);
  url.searchParams.set("client_id", cfg.clientId);
  url.searchParams.set("redirect_uri", cfg.redirectUri);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("scope", "core");
  url.searchParams.set("state", state);
  url.searchParams.set("code_challenge", codeChallenge);
  url.searchParams.set("code_challenge_method", "S256");
  return url.toString();
}

async function postTokenRequest(
  body: URLSearchParams,
  config: IgnitionConfig,
): Promise<OAuthTokenResponse> {
  const response = await fetch(config.tokenUrl, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: body.toString(),
    cache: "no-store",
  });

  const payload = (await response.json().catch(() => ({}))) as OAuthTokenResponse & {
    error?: string;
    error_description?: string;
    message?: string;
  };

  if (!response.ok || !payload.access_token) {
    const detail =
      payload.error_description ||
      payload.error ||
      payload.message ||
      `Token request failed (${response.status}).`;
    throw new Error(detail);
  }

  return payload;
}

export function tokenResponseToStore(
  payload: OAuthTokenResponse,
  previous?: IgnitionTokenSet | null,
): IgnitionTokenSet {
  const expiresIn =
    typeof payload.expires_in === "number" && Number.isFinite(payload.expires_in)
      ? payload.expires_in
      : 3600;
  const now = Date.now();

  return {
    accessToken: payload.access_token,
    refreshToken: payload.refresh_token || previous?.refreshToken || "",
    tokenType: payload.token_type || previous?.tokenType || "Bearer",
    scope: payload.scope || previous?.scope || "",
    expiresAt: new Date(now + expiresIn * 1000).toISOString(),
    obtainedAt: new Date(now).toISOString(),
  };
}

export async function exchangeAuthorizationCode(
  code: string,
  codeVerifier: string,
): Promise<IgnitionTokenSet> {
  const config = getIgnitionConfig();
  assertIgnitionConfigured(config);

  if (!codeVerifier) {
    throw new Error(
      "Missing PKCE code verifier. Start Connect to Ignition again.",
    );
  }

  const body = new URLSearchParams({
    grant_type: "authorization_code",
    code,
    redirect_uri: config.redirectUri,
    client_id: config.clientId,
    client_secret: config.clientSecret,
    code_verifier: codeVerifier,
  });

  const payload = await postTokenRequest(body, config);
  return tokenResponseToStore(payload);
}

export async function refreshAccessToken(
  refreshToken: string,
  previous?: IgnitionTokenSet | null,
): Promise<IgnitionTokenSet> {
  const config = getIgnitionConfig();
  assertIgnitionConfigured(config);

  if (!refreshToken) {
    throw new Error("No Ignition refresh token is stored. Reconnect Ignition.");
  }

  const body = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    client_id: config.clientId,
    client_secret: config.clientSecret,
  });

  const payload = await postTokenRequest(body, config);
  return tokenResponseToStore(payload, previous);
}
