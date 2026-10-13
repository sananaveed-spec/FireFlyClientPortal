"use client";

import { useMsal } from "@azure/msal-react";
import { getAccountEmail } from "@/auth/organization";
import { logoutCompletely } from "@/auth/session";

export function AuthHeader() {
  const { instance, accounts } = useMsal();
  const email = getAccountEmail(accounts[0]);
  const displayName = accounts[0]?.name ?? email;

  function handleLogout() {
    void logoutCompletely(instance);
  }

  if (!email) {
    return null;
  }

  return (
    <header className="auth-header">
      <div className="auth-header-text">
        <p className="auth-header-title">Fire Fly Client Portal</p>
        <p className="auth-header-meta">{displayName}</p>
      </div>
      <button type="button" className="button secondary" onClick={handleLogout}>
        Log out
      </button>
    </header>
  );
}
