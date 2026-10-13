"use client";

import { useMsal } from "@azure/msal-react";
import { allowedEmailDomainsLabel } from "@/auth/organization";

type UnauthorizedPageProps = {
  reason?: "domain" | "notListed";
};

export function UnauthorizedPage({ reason = "domain" }: UnauthorizedPageProps) {
  const { accounts } = useMsal();
  const email = accounts[0]?.username;

  return (
    <div className="unauthorized-panel">
      <h2>Access restricted</h2>
      {reason === "notListed" ? (
        <p>
          Your account is signed in, but it is not on the User Management list.
          Ask an Admin to add <strong>{email || "your email"}</strong> as Admin
          or Team Member.
        </p>
      ) : (
        <p>
          This application is only available to{" "}
          <strong>{allowedEmailDomainsLabel}</strong> accounts.
        </p>
      )}
      {email ? <p className="unauthorized-email">Signed in as: {email}</p> : null}
      <p className="unauthorized-hint">Use Log out above to switch accounts.</p>
    </div>
  );
}
