import { copyFile, mkdir, readFile, unlink, writeFile } from "fs/promises";
import path from "path";
import { getDataDir } from "@/lib/data-store";
import { getIgnitionConfig } from "@/lib/ignition/config";
import { refreshAccessToken } from "@/lib/ignition/oauth";

export type IgnitionTokenSet = {
  accessToken: string;
  refreshToken: string;
  tokenType: string;
  scope: string;
  expiresAt: string;
  obtainedAt: string;
};

const STATE_COOKIE = "ignition_oauth_state";
const PKCE_COOKIE = "ignition_oauth_pkce";
const REFRESH_SKEW_MS = 60_000;

function tokensPath() {
  return path.join(getDataDir(), "ignition-tokens.json");
}

async function ensureDataDir() {
  await mkdir(getDataDir(), { recursive: true });
}

async function writeJsonFile(filePath: string, value: unknown) {
  const tempPath = `${filePath}.${process.pid}.tmp`;
  const payload = `${JSON.stringify(value, null, 2)}\n`;
  await writeFile(tempPath, payload, "utf8");
  await copyFile(tempPath, filePath);
  await unlink(tempPath).catch(() => undefined);
}

function isMissingFile(error: unknown) {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    (error as { code?: string }).code === "ENOENT"
  );
}

export function getOAuthStateCookieName() {
  return STATE_COOKIE;
}

export function getOAuthPkceCookieName() {
  return PKCE_COOKIE;
}

export async function readIgnitionTokens(): Promise<IgnitionTokenSet | null> {
  await ensureDataDir();
  try {
    const raw = await readFile(tokensPath(), "utf8");
    const parsed = JSON.parse(raw.replace(/^\uFEFF/, "")) as Partial<IgnitionTokenSet>;
    if (
      typeof parsed !== "object" ||
      parsed === null ||
      typeof parsed.accessToken !== "string" ||
      !parsed.accessToken
    ) {
      return null;
    }
    return {
      accessToken: parsed.accessToken,
      refreshToken:
        typeof parsed.refreshToken === "string" ? parsed.refreshToken : "",
      tokenType:
        typeof parsed.tokenType === "string" && parsed.tokenType
          ? parsed.tokenType
          : "Bearer",
      scope: typeof parsed.scope === "string" ? parsed.scope : "",
      expiresAt:
        typeof parsed.expiresAt === "string" ? parsed.expiresAt : "",
      obtainedAt:
        typeof parsed.obtainedAt === "string" ? parsed.obtainedAt : "",
    };
  } catch (error) {
    if (isMissingFile(error)) {
      return null;
    }
    throw error;
  }
}

export async function writeIgnitionTokens(
  tokens: IgnitionTokenSet,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(tokensPath(), tokens);
}

export async function clearIgnitionTokens(): Promise<void> {
  await ensureDataDir();
  await unlink(tokensPath()).catch((error: unknown) => {
    if (!isMissingFile(error)) {
      throw error;
    }
  });
}

export function isTokenExpired(
  tokens: IgnitionTokenSet,
  skewMs = REFRESH_SKEW_MS,
): boolean {
  if (!tokens.expiresAt) {
    return false;
  }
  const expiresAt = Date.parse(tokens.expiresAt);
  if (!Number.isFinite(expiresAt)) {
    return false;
  }
  return Date.now() >= expiresAt - skewMs;
}

export async function getValidAccessToken(): Promise<{
  accessToken: string;
  tokens: IgnitionTokenSet;
}> {
  let tokens = await readIgnitionTokens();
  if (!tokens?.accessToken) {
    throw new Error("Ignition is not connected. Connect Ignition first.");
  }

  if (isTokenExpired(tokens)) {
    if (!tokens.refreshToken) {
      throw new Error(
        "Ignition access token expired and no refresh token is available. Reconnect Ignition.",
      );
    }
    tokens = await refreshAccessToken(tokens.refreshToken, tokens);
    await writeIgnitionTokens(tokens);
  }

  return { accessToken: tokens.accessToken, tokens };
}

export function ignitionConnectionStatus(tokens: IgnitionTokenSet | null): {
  connected: boolean;
  expiresAt: string | null;
  configured: boolean;
} {
  const config = getIgnitionConfig();
  return {
    connected: Boolean(tokens?.accessToken),
    expiresAt: tokens?.expiresAt || null,
    configured: Boolean(config.clientId && config.clientSecret),
  };
}
