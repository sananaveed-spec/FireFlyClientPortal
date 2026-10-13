export const azureClientId = process.env.NEXT_PUBLIC_AZURE_CLIENT_ID ?? "";
export const azureTenantId = process.env.NEXT_PUBLIC_AZURE_TENANT_ID ?? "";

export const isAzureConfigured =
  azureClientId.length > 0 &&
  azureTenantId.length > 0 &&
  azureClientId !== "your-application-client-id" &&
  azureTenantId !== "your-directory-tenant-id";
