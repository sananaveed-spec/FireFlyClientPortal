import { cookies } from "next/headers";
import { NextResponse } from "next/server";
import { getIgnitionPublicBaseUrl } from "@/lib/ignition/config";
import { exchangeAuthorizationCode } from "@/lib/ignition/oauth";
import {
  getOAuthPkceCookieName,
  getOAuthStateCookieName,
  writeIgnitionTokens,
} from "@/lib/ignition/tokens";

function portalHomeBase(request: Request): string {
  // Prefer public portal URL — request.url is often https://0.0.0.0:3000 in Docker.
  const publicBase = getIgnitionPublicBaseUrl().replace(/\/+$/, "");
  if (publicBase) {
    return publicBase;
  }
  const fromRequest = new URL(request.url);
  if (/0\.0\.0\.0/i.test(fromRequest.hostname)) {
    return "https://clientportal.fireflyaccountancy.com";
  }
  return fromRequest.origin;
}

function portalRedirect(request: Request, query: Record<string, string>) {
  const url = new URL("/", portalHomeBase(request));
  for (const [key, value] of Object.entries(query)) {
    url.searchParams.set(key, value);
  }
  return NextResponse.redirect(url);
}

function clearOAuthCookies(response: NextResponse) {
  response.cookies.set(getOAuthStateCookieName(), "", { path: "/", maxAge: 0 });
  response.cookies.set(getOAuthPkceCookieName(), "", { path: "/", maxAge: 0 });
  return response;
}

export async function GET(request: Request) {
  const url = new URL(request.url);
  const code = url.searchParams.get("code")?.trim() ?? "";
  const state = url.searchParams.get("state")?.trim() ?? "";
  const oauthError =
    url.searchParams.get("error_description") ||
    url.searchParams.get("error") ||
    "";

  const stateCookieName = getOAuthStateCookieName();
  const pkceCookieName = getOAuthPkceCookieName();
  const cookieStore = await cookies();
  const expectedState = cookieStore.get(stateCookieName)?.value ?? "";
  const codeVerifier = cookieStore.get(pkceCookieName)?.value ?? "";

  if (oauthError) {
    return clearOAuthCookies(
      portalRedirect(request, {
        tab: "ignition",
        ignition: "error",
        message: oauthError,
      }),
    );
  }

  if (!code) {
    return clearOAuthCookies(
      portalRedirect(request, {
        tab: "ignition",
        ignition: "error",
        message: "Ignition did not return an authorization code.",
      }),
    );
  }

  if (!state || !expectedState || state !== expectedState) {
    return clearOAuthCookies(
      portalRedirect(request, {
        tab: "ignition",
        ignition: "error",
        message: "Ignition OAuth state mismatch. Try connecting again.",
      }),
    );
  }

  if (!codeVerifier) {
    return clearOAuthCookies(
      portalRedirect(request, {
        tab: "ignition",
        ignition: "error",
        message: "Missing PKCE verifier. Start Connect to Ignition again.",
      }),
    );
  }

  try {
    const tokens = await exchangeAuthorizationCode(code, codeVerifier);
    await writeIgnitionTokens(tokens);

    return clearOAuthCookies(
      portalRedirect(request, {
        tab: "ignition",
        ignition: "connected",
      }),
    );
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to complete Ignition connection.";
    return clearOAuthCookies(
      portalRedirect(request, {
        tab: "ignition",
        ignition: "error",
        message,
      }),
    );
  }
}
