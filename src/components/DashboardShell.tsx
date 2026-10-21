"use client";

import { useEffect, useState } from "react";
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

type DashboardShellProps = {
  userName?: string | null;
};

function tabFromSearch(): DashboardTab | null {
  if (typeof window === "undefined") {
    return null;
  }
  const tab = new URLSearchParams(window.location.search).get("tab");
  if (tab === "ignition") {
    return "ignition";
  }
  return null;
}

export function DashboardShell({ userName }: DashboardShellProps) {
  const { canManageUsers, canEdit } = useAccess();
  const [activeTab, setActiveTab] = useState<DashboardTab>("clientProfile");

  useEffect(() => {
    const fromUrl = tabFromSearch();
    if (fromUrl === "ignition" && canManageUsers) {
      setActiveTab("ignition");
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
