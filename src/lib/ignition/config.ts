export type IgnitionConfig = {
  clientId: string;
  clientSecret: string;
  redirectUri: string;
  authUrl: string;
  tokenUrl: string;
  apiBase: string;
};

function trimEnv(name: string): string {
  return process.env[name]?.trim() ?? "";
}

export function getIgnitionConfig(): IgnitionConfig {
  const clientId = trimEnv("IGNITION_CLIENT_ID");
  const clientSecret = trimEnv("IGNITION_CLIENT_SECRET");
  const redirectUri =
    trimEnv("IGNITION_REDIRECT_URI") ||
    "http://localhost:3000/api/ignition/callback";
  const authUrl =
    trimEnv("IGNITION_AUTH_URL") ||
    "https://developers.ignitionapp.com/oauth2/authorize";
  const tokenUrl =
    trimEnv("IGNITION_TOKEN_URL") ||
    "https://developers.ignitionapp.com/oauth2/token";
  const apiBase =
    trimEnv("IGNITION_API_BASE") || "https://api.ignitionapp.com/rest/v1";

  return {
    clientId,
    clientSecret,
    redirectUri,
    authUrl,
    tokenUrl,
    apiBase: apiBase.replace(/\/+$/, ""),
  };
}

export function assertIgnitionConfigured(config: IgnitionConfig): void {
  if (!config.clientId || !config.clientSecret) {
    throw new Error(
      "Ignition is not configured. Set IGNITION_CLIENT_ID and IGNITION_CLIENT_SECRET in the server environment.",
    );
  }
  if (!config.redirectUri) {
    throw new Error("IGNITION_REDIRECT_URI is not configured.");
  }
}
