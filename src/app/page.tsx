"use client";

import { useIsAuthenticated, useMsal } from "@azure/msal-react";
import { useEffect, useState } from "react";
import { resolveAccess, type ResolvedAccess } from "@/auth/access";
import { AccessProvider } from "@/auth/accessContext";
import { AuthHeader } from "@/components/AuthHeader";
import { DashboardShell } from "@/components/DashboardShell";
import { LoginPage } from "@/components/LoginPage";
import { UnauthorizedPage } from "@/components/UnauthorizedPage";
import {
  getAccountEmail,
  isAllowedOrganizationEmail,
} from "@/auth/organization";
import {
  readLocalAuthSession,
  type LocalAuthSession,
} from "@/lib/local-auth";
import type { PortalUser } from "@/lib/portal-users";

export default function Home() {
  const isAuthenticated = useIsAuthenticated();
  const { accounts } = useMsal();
  const authenticatedEmail = isAuthenticated
    ? getAccountEmail(accounts[0])
    : "";
  const onDomain =
    isAuthenticated && isAllowedOrganizationEmail(authenticatedEmail);

  const [localSession, setLocalSession] = useState<LocalAuthSession | null>(
    null,
  );
  const [localReady, setLocalReady] = useState(false);
  const [access, setAccess] = useState<ResolvedAccess | null>(null);
  const [accessReady, setAccessReady] = useState(false);

  useEffect(() => {
    setLocalSession(readLocalAuthSession());
    setLocalReady(true);
  }, []);

  const isLocalAuthenticated = Boolean(localSession);
  const effectiveEmail = isLocalAuthenticated
    ? localSession!.email
    : authenticatedEmail;
  const displayName = isLocalAuthenticated
    ? localSession!.name || localSession!.email
    : (accounts[0]?.name ?? authenticatedEmail);

  useEffect(() => {
    let cancelled = false;

    async function loadAccess() {
      if (isLocalAuthenticated && effectiveEmail) {
        if (!cancelled) {
          setAccess(resolveAccess(effectiveEmail, []));
          setAccessReady(true);
        }
        return;
      }

      if (!onDomain || !authenticatedEmail) {
        if (!cancelled) {
          setAccess(null);
          setAccessReady(true);
        }
        return;
      }

      setAccessReady(false);
      try {
        const response = await fetch("/api/portal-users", { cache: "no-store" });
        const data = (await response.json()) as {
          users?: PortalUser[];
          error?: string;
        };
        const users = Array.isArray(data.users) ? data.users : [];
        if (!cancelled) {
          setAccess(resolveAccess(authenticatedEmail, users));
        }
      } catch {
        if (!cancelled) {
          setAccess(resolveAccess(authenticatedEmail, []));
        }
      } finally {
        if (!cancelled) {
          setAccessReady(true);
        }
      }
    }

    void loadAccess();
    return () => {
      cancelled = true;
    };
  }, [
    isLocalAuthenticated,
    effectiveEmail,
    onDomain,
    authenticatedEmail,
  ]);

  if (!localReady || ((isAuthenticated && onDomain) || isLocalAuthenticated) && !accessReady) {
    return (
      <main className="page">
        <div className="layout">
          <div className="card">
            <p className="table-empty">Checking access…</p>
          </div>
        </div>
      </main>
    );
  }

  if (access?.canView) {
    return (
      <AccessProvider access={access}>
        <DashboardShell userName={displayName} />
      </AccessProvider>
    );
  }

  return (
    <main className="page">
      <div className="layout">
        {isAuthenticated ? <AuthHeader /> : null}

        <div className="card">
          <div className="login-brand">
            <div className="brand-mark brand-mark--light" aria-hidden="true">
              <span className="brand-mark-glow" />
            </div>
            <h1 className="dashboard-title">Fire Fly Client Portal</h1>
          </div>

          {!isAuthenticated && !isLocalAuthenticated ? (
            <>
              <div className="divider" />
              <LoginPage />
            </>
          ) : (
            <>
              <div className="divider" />
              <UnauthorizedPage
                reason={
                  onDomain
                    ? "notListed"
                    : "domain"
                }
              />
            </>
          )}
        </div>
      </div>
    </main>
  );
}
