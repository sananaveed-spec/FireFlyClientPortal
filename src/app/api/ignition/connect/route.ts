import { randomBytes } from "crypto";
import { NextResponse } from "next/server";
import {
  assertIgnitionConfigured,
  getIgnitionConfig,
} from "@/lib/ignition/config";
import { buildAuthorizeUrl, createPkcePair } from "@/lib/ignition/oauth";
import {
  getOAuthPkceCookieName,
  getOAuthStateCookieName,
} from "@/lib/ignition/tokens";
import { requireAdmin } from "@/lib/requireEditor";

function oauthCookieOptions() {
  return {
    httpOnly: true,
    sameSite: "lax" as const,
    secure: process.env.NODE_ENV === "production",
    path: "/",
    maxAge: 60 * 10,
  };
}

export async function GET(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const config = getIgnitionConfig();
    assertIgnitionConfigured(config);

    const state = randomBytes(24).toString("hex");
    const { codeVerifier, codeChallenge } = createPkcePair();
    const authorizeUrl = buildAuthorizeUrl(state, codeChallenge, config);
    const response = NextResponse.json({ authorizeUrl });
    const cookieOptions = oauthCookieOptions();
    response.cookies.set(getOAuthStateCookieName(), state, cookieOptions);
    response.cookies.set(getOAuthPkceCookieName(), codeVerifier, cookieOptions);
    return response;
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to start Ignition connect.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
