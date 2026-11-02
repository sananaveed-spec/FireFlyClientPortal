import { azureClientId, azureTenantId } from "@/auth/env";

/**
 * App-only Graph token for server jobs (Ignition webhook → SharePoint).
 * Requires AZURE_CLIENT_SECRET and Application permissions with admin consent:
 * Group.ReadWrite.All, User.Invite.All, User.Read.All, Files.ReadWrite.All,
 * Directory.Read.All (or Sites.Read.All).
 */
export async function acquireGraphAppOnlyToken(): Promise<string> {
  const clientId = azureClientId.trim();
  const tenantId = azureTenantId.trim();
  const clientSecret = process.env.AZURE_CLIENT_SECRET?.trim() || "";

  if (!clientId || !tenantId) {
    throw new Error(
      "Azure AD is not configured (NEXT_PUBLIC_AZURE_CLIENT_ID / NEXT_PUBLIC_AZURE_TENANT_ID).",
    );
  }
  if (!clientSecret) {
    throw new Error(
      "AZURE_CLIENT_SECRET is not set. Add the app registration client secret so proposal.accepted can create SharePoint sites.",
    );
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const payload = (await response.json().catch(() => ({}))) as {
    access_token?: string;
    error?: string;
    error_description?: string;
  };
  if (!response.ok || !payload.access_token) {
    throw new Error(
      payload.error_description ||
        payload.error ||
        `Failed to acquire Graph app token (${response.status}).`,
    );
  }
  return payload.access_token;
}

export function isGraphAppOnlyConfigured(): boolean {
  return Boolean(
    azureClientId.trim() &&
      azureTenantId.trim() &&
      process.env.AZURE_CLIENT_SECRET?.trim(),
  );
}
