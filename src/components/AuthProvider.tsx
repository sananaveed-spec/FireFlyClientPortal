"use client";

import { MsalProvider } from "@azure/msal-react";
import { useEffect, useState, type ReactNode } from "react";
import { msalInstance } from "@/auth/msalInstance";

type AuthProviderProps = {
  children: ReactNode;
};

export function AuthProvider({ children }: AuthProviderProps) {
  const [ready, setReady] = useState(false);

  useEffect(() => {
    async function initializeAuth() {
      await msalInstance.initialize();

      try {
        const response = await msalInstance.handleRedirectPromise();
        if (response?.account) {
          msalInstance.setActiveAccount(response.account);
        }
      } catch {
        // No redirect in progress on a normal page load.
      }

      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0 && !msalInstance.getActiveAccount()) {
        msalInstance.setActiveAccount(accounts[0]);
      }

      setReady(true);
    }

    void initializeAuth();
  }, []);

  if (!ready) {
    return (
      <main className="page">
        <div className="layout">
          <div className="card">
            <p className="table-empty">Loading Fire Fly…</p>
          </div>
        </div>
      </main>
    );
  }

  return <MsalProvider instance={msalInstance}>{children}</MsalProvider>;
}
