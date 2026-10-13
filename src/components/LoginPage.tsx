"use client";

import { useMsal } from "@azure/msal-react";
import { isAzureConfigured } from "@/auth/env";
import { allowedEmailDomainsLabel } from "@/auth/organization";
import { loginWithCredentials } from "@/auth/session";

export function LoginPage() {
  const { instance } = useMsal();

  function handleSignIn() {
    if (!isAzureConfigured) {
      window.alert(
        "Azure AD is not configured. Add NEXT_PUBLIC_AZURE_CLIENT_ID and NEXT_PUBLIC_AZURE_TENANT_ID to .env.local, then restart npm run dev.",
      );
      return;
    }

    void loginWithCredentials(instance);
  }

  return (
    <div className="login-panel">
      <div className="microsoft-logo">
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="108"
          height="24"
          viewBox="0 0 108 24"
          aria-hidden="true"
        >
          <rect x="0" y="0" width="11" height="11" fill="#f25022" />
          <rect x="13" y="0" width="11" height="11" fill="#7fba00" />
          <rect x="0" y="13" width="11" height="11" fill="#00a4ef" />
          <rect x="13" y="13" width="11" height="11" fill="#ffb900" />
          <text
            x="32"
            y="17"
            fill="#5a6b80"
            fontFamily="Segoe UI, sans-serif"
            fontSize="18"
            fontWeight="600"
          >
            Microsoft
          </text>
        </svg>
      </div>

      <p className="login-subtitle">
        Log in with <strong>{allowedEmailDomainsLabel}</strong> account to
        access Fire Fly Client Portal.
      </p>

      {!isAzureConfigured ? (
        <div className="config-warning" role="alert">
          <strong>Azure AD is not configured.</strong>
          <p>
            Create a <code>.env.local</code> file in the project root with your
            Azure app registration values, then restart <code>npm run dev</code>.
          </p>
          <pre>{`NEXT_PUBLIC_AZURE_CLIENT_ID=your-client-id
NEXT_PUBLIC_AZURE_TENANT_ID=your-tenant-id`}</pre>
        </div>
      ) : null}

      <button
        type="button"
        className={`sign-in-button${isAzureConfigured ? "" : " sign-in-button--disabled"}`}
        onClick={handleSignIn}
      >
        Log in with Microsoft
      </button>
    </div>
  );
}
