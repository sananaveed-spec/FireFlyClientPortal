"use client";

import { useMsal } from "@azure/msal-react";
import { useState, type FormEvent } from "react";
import { isAzureConfigured } from "@/auth/env";
import { allowedEmailDomainsLabel } from "@/auth/organization";
import { loginWithCredentials } from "@/auth/session";
import {
  LOCAL_TEST_EMAIL,
  LOCAL_TEST_NAME,
  writeLocalAuthSession,
} from "@/lib/local-auth";

export function LoginPage() {
  const { instance } = useMsal();
  const [id, setId] = useState(LOCAL_TEST_EMAIL);
  const [password, setPassword] = useState("");
  const [localError, setLocalError] = useState<string | null>(null);
  const [localSaving, setLocalSaving] = useState(false);

  function handleSignIn() {
    if (!isAzureConfigured) {
      window.alert(
        "Azure AD is not configured. Add NEXT_PUBLIC_AZURE_CLIENT_ID and NEXT_PUBLIC_AZURE_TENANT_ID to .env.local, then restart npm run dev.",
      );
      return;
    }

    void loginWithCredentials(instance);
  }

  async function handleLocalSignIn(event: FormEvent) {
    event.preventDefault();
    setLocalError(null);
    setLocalSaving(true);
    try {
      const response = await fetch("/api/local-login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id, password }),
      });
      const data = (await response.json()) as {
        email?: string;
        name?: string;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Invalid ID or password.");
      }
      writeLocalAuthSession({
        email: data.email || LOCAL_TEST_EMAIL,
        name: data.name || LOCAL_TEST_NAME,
        loggedInAt: new Date().toISOString(),
      });
      window.location.assign(window.location.origin);
    } catch (error) {
      setLocalError(
        error instanceof Error ? error.message : "Invalid ID or password.",
      );
    } finally {
      setLocalSaving(false);
    }
  }

  return (
    <div className="login-panel">
      <form className="local-login-form" onSubmit={handleLocalSignIn}>
        <p className="login-subtitle">Log in with ID and password</p>
        <label className="field">
          <span className="field-label">ID</span>
          <input
            className="field-input"
            type="email"
            autoComplete="username"
            value={id}
            onChange={(event) => setId(event.target.value)}
            required
          />
        </label>
        <label className="field">
          <span className="field-label">Password</span>
          <input
            className="field-input"
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            required
          />
        </label>
        {localError ? <p className="form-message error">{localError}</p> : null}
        <button
          type="submit"
          className="sign-in-button"
          disabled={localSaving}
        >
          {localSaving ? "Signing in…" : "Log in with ID"}
        </button>
      </form>

      <div className="login-or" aria-hidden="true">
        <span>or</span>
      </div>

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
