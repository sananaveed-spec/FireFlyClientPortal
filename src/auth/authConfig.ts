import type { Configuration } from "@azure/msal-browser";
import { azureClientId, azureTenantId } from "./env";

/**
 * Prefer single-tenant authority so only your Azure AD directory can sign in.
 * Override with NEXT_PUBLIC_AZURE_AUTHORITY if needed.
 */
const azureAuthority =
  process.env.NEXT_PUBLIC_AZURE_AUTHORITY?.trim() ||
  (azureTenantId
    ? `https://login.microsoftonline.com/${azureTenantId}`
    : "https://login.microsoftonline.com/organizations");

export function getMsalConfig(): Configuration {
  return {
    auth: {
      clientId: azureClientId,
      authority: azureAuthority,
      redirectUri:
        typeof window !== "undefined" ? window.location.origin : "/",
      postLogoutRedirectUri:
        typeof window !== "undefined" ? window.location.origin : "/",
    },
    cache: {
      cacheLocation: "sessionStorage",
    },
  };
}

export const loginRequest = {
  scopes: ["User.Read", "Mail.Send"],
};

/** Use when connecting a mailbox outside the Fire Fly home tenant. */
export const organizationsAuthority =
  "https://login.microsoftonline.com/organizations";

/**
 * Token scopes for sending mail as the signed-in user via Microsoft Graph.
 * Uses /organizations so a work account like sana@epsfresno.com can consent
 * in its home tenant (app registration must allow multi-tenant accounts).
 */
export const mailSendRequest = {
  scopes: ["Mail.Send", "User.Read"],
  authority: organizationsAuthority,
};
