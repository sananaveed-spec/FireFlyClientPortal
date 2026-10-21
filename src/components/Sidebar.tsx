"use client";

import { useEffect, useId, useRef, useState } from "react";
import { logoutCompletely } from "@/auth/session";
import { useMsal } from "@azure/msal-react";

export type DashboardTab =
  | "clientProfile"
  | "resourceAllocation"
  | "clientManagement"
  | "lookupData"
  | "history"
  | "userManagement"
  | "ignition"
  | "profile";

const NAV_ITEMS: { id: DashboardTab; label: string; adminOnly?: boolean }[] = [
  { id: "clientProfile", label: "Client Profile" },
  { id: "resourceAllocation", label: "Resource Allocation" },
  { id: "clientManagement", label: "Client Management" },
  { id: "lookupData", label: "Lookup Data" },
  { id: "history", label: "History" },
  { id: "userManagement", label: "User Management", adminOnly: true },
  { id: "ignition", label: "Ignition", adminOnly: true },
];

type SidebarProps = {
  activeTab: DashboardTab;
  onTabChange: (tab: DashboardTab) => void;
  userName?: string | null;
  canManageUsers?: boolean;
};

function getInitials(userName?: string | null) {
  const parts = String(userName ?? "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  if (parts.length === 0) {
    return "U";
  }
  if (parts.length === 1) {
    return parts[0]!.slice(0, 2).toUpperCase();
  }
  return `${parts[0]![0] ?? ""}${parts[1]![0] ?? ""}`.toUpperCase();
}

export function Sidebar({
  activeTab,
  onTabChange,
  userName,
  canManageUsers = false,
}: SidebarProps) {
  const { instance } = useMsal();
  const [menuOpen, setMenuOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement | null>(null);
  const menuId = useId();
  const initials = getInitials(userName);

  useEffect(() => {
    if (!menuOpen) {
      return;
    }

    function onPointerDown(event: globalThis.MouseEvent) {
      if (!menuRef.current?.contains(event.target as Node)) {
        setMenuOpen(false);
      }
    }

    function onKeyDown(event: globalThis.KeyboardEvent) {
      if (event.key === "Escape") {
        setMenuOpen(false);
      }
    }

    document.addEventListener("mousedown", onPointerDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("mousedown", onPointerDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [menuOpen]);

  function handleSignOut() {
    setMenuOpen(false);
    void logoutCompletely(instance);
  }

  function openProfile() {
    setMenuOpen(false);
    onTabChange("profile");
  }

  const visibleItems = NAV_ITEMS.filter(
    (item) => !item.adminOnly || canManageUsers,
  );

  return (
    <aside className="sidebar">
      <div className="brand-block">
        <img
          className="brand-logo"
          src="/firefly-logo.webp"
          alt="Fire Fly Accountancy"
          width={180}
          height={48}
        />

        <div className="account-menu" ref={menuRef}>
          <button
            type="button"
            className={
              menuOpen || activeTab === "profile"
                ? "account-avatar account-avatar--active"
                : "account-avatar"
            }
            aria-haspopup="menu"
            aria-expanded={menuOpen}
            aria-controls={menuId}
            title={userName ?? "Account"}
            onClick={() => setMenuOpen((open) => !open)}
          >
            <span className="account-avatar-initials">{initials}</span>
          </button>

          {menuOpen ? (
            <div
              id={menuId}
              className="account-dropdown"
              role="menu"
              aria-label="Manage Account"
            >
              <p className="account-dropdown-heading">Manage Account</p>
              <button
                type="button"
                className="account-dropdown-item"
                role="menuitem"
                onClick={openProfile}
              >
                <span className="account-dropdown-icon" aria-hidden="true">
                  <svg viewBox="0 0 24 24" width="18" height="18" fill="none">
                    <circle
                      cx="12"
                      cy="8"
                      r="3.25"
                      stroke="currentColor"
                      strokeWidth="1.7"
                    />
                    <path
                      d="M5.5 18.25c1.6-2.7 3.7-4 6.5-4s4.9 1.3 6.5 4"
                      stroke="currentColor"
                      strokeWidth="1.7"
                      strokeLinecap="round"
                    />
                  </svg>
                </span>
                My Profile
              </button>
              <button
                type="button"
                className="account-dropdown-item"
                role="menuitem"
                onClick={handleSignOut}
              >
                <span className="account-dropdown-icon" aria-hidden="true">
                  <svg viewBox="0 0 24 24" width="18" height="18" fill="none">
                    <path
                      d="M10 5.5H6.75A1.75 1.75 0 0 0 5 7.25v9.5c0 .97.78 1.75 1.75 1.75H10"
                      stroke="currentColor"
                      strokeWidth="1.7"
                      strokeLinecap="round"
                    />
                    <path
                      d="M10.5 12H19"
                      stroke="currentColor"
                      strokeWidth="1.7"
                      strokeLinecap="round"
                    />
                    <path
                      d="M16.25 8.75 19.5 12l-3.25 3.25"
                      stroke="currentColor"
                      strokeWidth="1.7"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                </span>
                Sign out
              </button>
            </div>
          ) : null}
        </div>
      </div>

      <nav className="nav-list" aria-label="Fire Fly Client Portal">
        {visibleItems.map((item) => (
          <button
            key={item.id}
            type="button"
            className={
              activeTab === item.id ? "nav-link active" : "nav-link"
            }
            onClick={() => onTabChange(item.id)}
          >
            {item.label}
          </button>
        ))}
      </nav>
    </aside>
  );
}
