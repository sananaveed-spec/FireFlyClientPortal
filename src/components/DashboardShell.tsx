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
  pendingAddId?: string;
  action?: string;
  createdAt: string;
};

type PortalChoice = "firefly-sharepoint" | "other" | "skip";

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
  const [addTarget, setAddTarget] = useState<DashboardNotificationRow | null>(
    null,
  );
  const [portalChoice, setPortalChoice] =
    useState<PortalChoice>("firefly-sharepoint");
  const [otherPortalLink, setOtherPortalLink] = useState("");
  const [addBusy, setAddBusy] = useState(false);
  const [addError, setAddError] = useState<string | null>(null);

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

  const openAddToSystem = useCallback((note: DashboardNotificationRow) => {
    setAddTarget(note);
    setPortalChoice("firefly-sharepoint");
    setOtherPortalLink("");
    setAddError(null);
  }, []);

  const closeAddToSystem = useCallback(() => {
    if (addBusy) {
      return;
    }
    setAddTarget(null);
    setAddError(null);
  }, [addBusy]);

  const submitAddToSystem = useCallback(async () => {
    if (!addTarget?.pendingAddId) {
      setAddError("Missing pending add id.");
      return;
    }
    setAddBusy(true);
    setAddError(null);
    try {
      const response = await portalFetch("/api/ignition/pending-adds", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "add",
          pendingId: addTarget.pendingAddId,
          portalChoice,
          portalLink: portalChoice === "other" ? otherPortalLink.trim() : "",
        }),
      });
      const data = (await response.json()) as {
        error?: string;
        legalName?: string;
        portalNotes?: string;
      };
      if (!response.ok) {
        setAddError(data.error || "Could not add client to the system.");
        return;
      }
      setAddTarget(null);
      setActiveTab("clientProfile");
      await loadNotifications();
      window.dispatchEvent(new Event("dashboard-notifications-refresh"));
    } catch {
      setAddError("Could not add client to the system.");
    } finally {
      setAddBusy(false);
    }
  }, [
    addTarget,
    portalChoice,
    otherPortalLink,
    loadNotifications,
  ]);

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
              {notifications.slice(0, 8).map((note) => {
                const needsAdd =
                  note.action === "add_to_system" &&
                  Boolean(note.pendingAddId);
                return (
                  <div
                    key={note.id}
                    className={`inline-notice ${
                      needsAdd
                        ? "inline-notice--warning"
                        : "inline-notice--success"
                    }`}
                    role="status"
                  >
                    <p className="inline-notice__text">
                      <strong>{note.title}</strong>
                      {note.message ? ` — ${note.message}` : ""}
                      {needsAdd && canManageUsers ? (
                        <>
                          {" "}
                          <button
                            type="button"
                            className="table-link-button"
                            onClick={() => openAddToSystem(note)}
                          >
                            Add to system
                          </button>
                        </>
                      ) : null}
                      {note.clientProfileId && !needsAdd ? (
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
                );
              })}
            </div>
          ) : null}

          {addTarget ? (
            <div
              className="dialog-backdrop"
              role="presentation"
              onClick={closeAddToSystem}
            >
              <div
                className="dialog-panel dialog-panel--add-to-system"
                role="dialog"
                aria-modal="true"
                aria-labelledby="ignition-add-title"
                onClick={(event) => event.stopPropagation()}
              >
                <div className="dialog-header">
                  <div className="add-to-system-heading">
                    <p className="add-to-system-eyebrow">Proposal accepted</p>
                    <h2 id="ignition-add-title" className="dialog-title">
                      Add to system
                    </h2>
                  </div>
                  <button
                    type="button"
                    className="dialog-close"
                    aria-label="Close"
                    disabled={addBusy}
                    onClick={closeAddToSystem}
                  >
                    ×
                  </button>
                </div>

                <div className="add-to-system-client">
                  <span className="add-to-system-client-label">Client</span>
                  <strong className="add-to-system-client-name">
                    {addTarget.clientName || "Client"}
                  </strong>
                </div>

                <p className="add-to-system-lead">
                  Finish setup by choosing how this client’s document portal
                  should be linked. You can change the portal later in Client
                  Info.
                </p>

                <fieldset className="add-to-system-choices">
                  <legend className="add-to-system-choices-legend">
                    Client document portal
                  </legend>

                  <label
                    className={
                      portalChoice === "firefly-sharepoint"
                        ? "add-to-system-choice add-to-system-choice--selected"
                        : "add-to-system-choice"
                    }
                  >
                    <input
                      type="radio"
                      name="ignition-portal-choice"
                      checked={portalChoice === "firefly-sharepoint"}
                      disabled={addBusy}
                      onChange={() => setPortalChoice("firefly-sharepoint")}
                    />
                    <span className="add-to-system-choice-body">
                      <span className="add-to-system-choice-title">
                        FireFly SharePoint
                      </span>
                      <span className="add-to-system-choice-desc">
                        Create a Fire Fly SharePoint site and save it as the
                        Portal Link. Recommended for most clients.
                      </span>
                    </span>
                  </label>

                  <label
                    className={
                      portalChoice === "other"
                        ? "add-to-system-choice add-to-system-choice--selected"
                        : "add-to-system-choice"
                    }
                  >
                    <input
                      type="radio"
                      name="ignition-portal-choice"
                      checked={portalChoice === "other"}
                      disabled={addBusy}
                      onChange={() => setPortalChoice("other")}
                    />
                    <span className="add-to-system-choice-body">
                      <span className="add-to-system-choice-title">Other</span>
                      <span className="add-to-system-choice-desc">
                        Use an existing portal URL (SharePoint, Dropbox, etc.).
                      </span>
                    </span>
                  </label>

                  {portalChoice === "other" ? (
                    <label className="field add-to-system-other-link">
                      <span className="field-label">
                        Portal link{" "}
                        <span className="add-to-system-optional">
                          (optional)
                        </span>
                      </span>
                      <input
                        className="field-input"
                        type="url"
                        placeholder="https://…"
                        value={otherPortalLink}
                        disabled={addBusy}
                        autoFocus
                        onChange={(event) =>
                          setOtherPortalLink(event.target.value)
                        }
                      />
                    </label>
                  ) : null}

                  <label
                    className={
                      portalChoice === "skip"
                        ? "add-to-system-choice add-to-system-choice--selected"
                        : "add-to-system-choice"
                    }
                  >
                    <input
                      type="radio"
                      name="ignition-portal-choice"
                      checked={portalChoice === "skip"}
                      disabled={addBusy}
                      onChange={() => setPortalChoice("skip")}
                    />
                    <span className="add-to-system-choice-body">
                      <span className="add-to-system-choice-title">
                        Skip for now
                      </span>
                      <span className="add-to-system-choice-desc">
                        Add Client Info without a portal link. Set it later
                        when ready.
                      </span>
                    </span>
                  </label>
                </fieldset>

                {addError ? (
                  <p className="form-message error" role="alert">
                    {addError}
                  </p>
                ) : null}

                <div className="dialog-actions add-to-system-actions">
                  <button
                    type="button"
                    className="button secondary"
                    disabled={addBusy}
                    onClick={closeAddToSystem}
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    className="button primary"
                    disabled={addBusy}
                    onClick={() => void submitAddToSystem()}
                  >
                    {addBusy
                      ? "Adding…"
                      : portalChoice === "skip"
                        ? "Add without portal"
                        : "Add client"}
                  </button>
                </div>
              </div>
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
