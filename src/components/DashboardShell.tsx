"use client";

import { useCallback, useEffect, useState } from "react";
import { useAccess } from "@/auth/accessContext";
import { ClientManagementPanel } from "@/components/ClientManagementPanel";
import { ClientProfilePanel } from "@/components/ClientProfilePanel";
import { HistoryPanel } from "@/components/HistoryPanel";
import { IgnitionPanel } from "@/components/IgnitionPanel";
import { LookupDataPanel } from "@/components/LookupDataPanel";
import { ProfilePanel } from "@/components/ProfilePanel";
import { ResourceAllocationPanel } from "@/components/ResourceAllocationPanel";
import { Sidebar, type DashboardTab } from "@/components/Sidebar";
import { UserManagementPanel } from "@/components/UserManagementPanel";
import { portalFetch } from "@/lib/portal-api";

type DashboardShellProps = {
  userName?: string | null;
};

type DashboardNotificationRow = {
  id: string;
  title: string;
  message: string;
  clientName: string;
  clientProfileId: string;
  createdAt: string;
};

function tabFromSearch(): DashboardTab | null {
  if (typeof window === "undefined") {
    return null;
  }
  const tab = new URLSearchParams(window.location.search).get("tab");
  if (tab === "ignition") {
    return "ignition";
  }
  if (tab === "clientProfile") {
    return "clientProfile";
  }
  return null;
}

export function DashboardShell({ userName }: DashboardShellProps) {
  const { canManageUsers, canEdit } = useAccess();
  const [activeTab, setActiveTab] = useState<DashboardTab>("clientProfile");
  const [notifications, setNotifications] = useState<
    DashboardNotificationRow[]
  >([]);

  const loadNotifications = useCallback(async () => {
    try {
      const response = await portalFetch("/api/dashboard-notifications", {
        cache: "no-store",
      });
      const data = (await response.json()) as {
        notifications?: DashboardNotificationRow[];
      };
      if (!response.ok) {
        return;
      }
      setNotifications(
        Array.isArray(data.notifications) ? data.notifications : [],
      );
    } catch {
      // Banner stays empty if notifications cannot load.
    }
  }, []);

  const dismissNotification = useCallback(async (id: string) => {
    setNotifications((current) => current.filter((item) => item.id !== id));
    try {
      const response = await portalFetch("/api/dashboard-notifications", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "dismiss", id }),
      });
      const data = (await response.json()) as {
        notifications?: DashboardNotificationRow[];
      };
      if (response.ok && Array.isArray(data.notifications)) {
        setNotifications(data.notifications);
      }
    } catch {
      // Optimistic dismiss already applied.
    }
  }, []);

  useEffect(() => {
    const fromUrl = tabFromSearch();
    if (fromUrl === "ignition" && canManageUsers) {
      setActiveTab("ignition");
    } else if (fromUrl === "clientProfile") {
      setActiveTab("clientProfile");
    }
  }, [canManageUsers]);

  useEffect(() => {
    if (
      (activeTab === "userManagement" || activeTab === "ignition") &&
      !canManageUsers
    ) {
      setActiveTab("clientProfile");
    }
  }, [activeTab, canManageUsers]);

  useEffect(() => {
    void loadNotifications();
    const timer = window.setInterval(() => {
      void loadNotifications();
    }, 60_000);
    const onRefresh = () => {
      void loadNotifications();
    };
    window.addEventListener("dashboard-notifications-refresh", onRefresh);
    return () => {
      window.clearInterval(timer);
      window.removeEventListener("dashboard-notifications-refresh", onRefresh);
    };
  }, [loadNotifications]);

  return (
    <div className="app-shell">
      <Sidebar
        activeTab={activeTab}
        onTabChange={setActiveTab}
        userName={userName}
        canManageUsers={canManageUsers}
      />
      <main className="main-pane">
        <div className="dashboard-main">
          <header className="page-header">
            <div>
              <h1>Welcome to Fire Fly Client Portal</h1>
              {!canEdit ? (
                <p className="welcome-subtitle">
                  View-only access (Team Member). Ask an Admin to change your
                  role if you need edit permissions.
                </p>
              ) : null}
            </div>
          </header>

          {notifications.length > 0 ? (
            <div className="dashboard-notifications" aria-live="polite">
              {notifications.slice(0, 8).map((note) => (
                <div
                  key={note.id}
                  className="inline-notice inline-notice--success"
                  role="status"
                >
                  <p className="inline-notice__text">
                    <strong>{note.title}</strong>
                    {note.message ? ` — ${note.message}` : ""}
                    {note.clientProfileId ? (
                      <>
                        {" "}
                        <button
                          type="button"
                          className="table-link-button"
                          onClick={() => setActiveTab("clientProfile")}
                        >
                          Open Client Info
                        </button>
                      </>
                    ) : null}
                  </p>
                  <button
                    type="button"
                    className="inline-notice__close"
                    aria-label="Dismiss notification"
                    onClick={() => void dismissNotification(note.id)}
                  >
                    ×
                  </button>
                </div>
              ))}
            </div>
          ) : null}

          <div className="dashboard-body">
            {activeTab === "clientProfile" ? (
              <ClientProfilePanel />
            ) : activeTab === "resourceAllocation" ? (
              <ResourceAllocationPanel />
            ) : activeTab === "clientManagement" ? (
              <ClientManagementPanel />
            ) : activeTab === "lookupData" ? (
              <LookupDataPanel />
            ) : activeTab === "userManagement" ? (
              <UserManagementPanel />
            ) : activeTab === "ignition" ? (
              <IgnitionPanel />
            ) : activeTab === "profile" ? (
              <ProfilePanel />
            ) : (
              <HistoryPanel />
            )}
          </div>
        </div>
      </main>
    </div>
  );
}
