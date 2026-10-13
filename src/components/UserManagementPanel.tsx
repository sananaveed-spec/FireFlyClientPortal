"use client";

import {
  useCallback,
  useEffect,
  useState,
  type FormEvent,
} from "react";
import { createPortal } from "react-dom";
import { useAccess } from "@/auth/accessContext";
import { portalFetch } from "@/lib/portal-api";
import {
  isValidEmail,
  normalizeEmail,
  notifyPortalUsersChanged,
  PORTAL_ROLES,
  PORTAL_USERS_CHANGED_EVENT,
  ROLE_LABELS,
  type PortalRole,
  type PortalUser,
} from "@/lib/portal-users";

type UsersResponse = {
  users?: PortalUser[];
  error?: string;
};

type EditorState = {
  mode: "add" | "edit";
  email: string;
  name: string;
  position: string;
  phone: string;
  role: PortalRole;
};

function emptyEditor(): EditorState {
  return {
    mode: "add",
    email: "",
    name: "",
    position: "",
    phone: "",
    role: "team",
  };
}

export function UserManagementPanel() {
  const { email: signedInEmail } = useAccess();
  const [users, setUsers] = useState<PortalUser[]>([]);
  const [selectedEmails, setSelectedEmails] = useState<string[]>([]);
  const [editor, setEditor] = useState<EditorState | null>(null);
  const [mounted, setMounted] = useState(false);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setMounted(true);
  }, []);

  const loadUsers = useCallback(async () => {
    setError(null);
    try {
      const response = await portalFetch("/api/portal-users", {
        cache: "no-store",
      });
      const data = (await response.json()) as UsersResponse;
      if (!response.ok) {
        throw new Error(data.error ?? "Failed to load users.");
      }
      const next = Array.isArray(data.users) ? data.users : [];
      setUsers(next);
      setSelectedEmails((current) =>
        current.filter((email) => next.some((user) => user.email === email)),
      );
    } catch (loadError) {
      setError(
        loadError instanceof Error ? loadError.message : "Failed to load users.",
      );
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadUsers();
  }, [loadUsers]);

  useEffect(() => {
    function onChanged() {
      void loadUsers();
    }
    window.addEventListener(PORTAL_USERS_CHANGED_EVENT, onChanged);
    return () => {
      window.removeEventListener(PORTAL_USERS_CHANGED_EVENT, onChanged);
    };
  }, [loadUsers]);

  const selectedUsers = users.filter((user) =>
    selectedEmails.includes(user.email),
  );
  const selectedUser =
    selectedUsers.length === 1 ? selectedUsers[0] ?? null : null;
  const adminCount = users.filter((user) => user.role === "admin").length;
  const selectedAdminCount = selectedUsers.filter(
    (user) => user.role === "admin",
  ).length;
  const wouldRemoveLastAdmin =
    selectedAdminCount > 0 && adminCount - selectedAdminCount < 1;
  const allSelected =
    users.length > 0 && selectedEmails.length === users.length;

  function toggleSelected(email: string) {
    setSelectedEmails((current) =>
      current.includes(email)
        ? current.filter((entry) => entry !== email)
        : [...current, email],
    );
  }

  function toggleSelectAll() {
    setSelectedEmails(allSelected ? [] : users.map((user) => user.email));
  }

  function openAddUser() {
    setError(null);
    setEditor(emptyEditor());
  }

  function openEditSelected() {
    if (!selectedUser) {
      return;
    }
    setError(null);
    setEditor({
      mode: "edit",
      email: selectedUser.email,
      name: selectedUser.name,
      position: selectedUser.position,
      phone: selectedUser.phone,
      role: selectedUser.role,
    });
  }

  function closeEditor() {
    if (saving) {
      return;
    }
    setEditor(null);
    setError(null);
  }

  async function saveUser(input: {
    email: string;
    role: PortalRole;
    name: string;
    position: string;
    phone: string;
  }): Promise<boolean> {
    setSaving(true);
    setError(null);
    try {
      const response = await portalFetch("/api/portal-users", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(input),
      });
      const data = (await response.json()) as UsersResponse;
      if (!response.ok) {
        throw new Error(data.error ?? "Failed to save user.");
      }
      const next = Array.isArray(data.users) ? data.users : [];
      setUsers(next);
      setSelectedEmails([input.email]);
      notifyPortalUsersChanged();
      return true;
    } catch (saveError) {
      setError(
        saveError instanceof Error ? saveError.message : "Failed to save user.",
      );
      return false;
    } finally {
      setSaving(false);
    }
  }

  async function submitEditor(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!editor) {
      return;
    }
    const email = normalizeEmail(editor.email);
    if (!isValidEmail(email)) {
      setError("Enter a valid email address.");
      return;
    }
    const ok = await saveUser({
      email,
      role: editor.role,
      name: editor.name.trim(),
      position: editor.position.trim(),
      phone: editor.phone.trim(),
    });
    if (ok) {
      setEditor(null);
    }
  }

  async function handleDeleteSelected() {
    if (selectedUsers.length === 0) {
      return;
    }
    if (wouldRemoveLastAdmin) {
      setError("At least one admin is required.");
      return;
    }

    const label =
      selectedUsers.length === 1
        ? selectedUsers[0]?.email
        : `${selectedUsers.length} users`;
    if (!window.confirm(`Remove ${label} from User Management?`)) {
      return;
    }

    setSaving(true);
    setError(null);
    try {
      let next = users;
      for (const user of selectedUsers) {
        const response = await portalFetch("/api/portal-users", {
          method: "DELETE",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ email: user.email }),
        });
        const data = (await response.json()) as UsersResponse;
        if (!response.ok) {
          throw new Error(data.error ?? "Failed to remove user.");
        }
        next = Array.isArray(data.users) ? data.users : next;
      }
      if (editor && selectedEmails.includes(editor.email)) {
        setEditor(null);
      }
      setUsers(next);
      setSelectedEmails([]);
      notifyPortalUsersChanged();
    } catch (deleteError) {
      setError(
        deleteError instanceof Error
          ? deleteError.message
          : "Failed to remove user.",
      );
    } finally {
      setSaving(false);
    }
  }

  return (
    <section className="content-panel content-panel--actions">
      <div className="panel-title-row users-header">
        <div className="users-header-copy">
          <h2 className="section-title">User Management</h2>
          <p className="lookup-hint">
            Admin can do anything. Team Member can view everything but cannot
            edit. Only listed @fireflyaccountancy.com users can open the portal.
          </p>
        </div>
        <div className="panel-title-actions">
          <button
            type="button"
            className="button danger"
            onClick={() => void handleDeleteSelected()}
            disabled={selectedUsers.length === 0 || saving}
          >
            Delete
          </button>
          <button
            type="button"
            className="button secondary-solid"
            onClick={openEditSelected}
            disabled={!selectedUser || saving}
          >
            Edit
          </button>
          <button
            type="button"
            className="button primary"
            onClick={openAddUser}
            disabled={saving}
          >
            Add
          </button>
        </div>
      </div>

      {error && !editor ? <p className="form-message error">{error}</p> : null}

      {loading ? (
        <p className="table-empty">Loading…</p>
      ) : users.length === 0 ? (
        <p className="table-empty">No users yet. Click Add to create one.</p>
      ) : (
        <div className="table-wrap">
          <table className="projects-table projects-table--with-check">
            <thead>
              <tr>
                <th className="col-check">
                  <input
                    type="checkbox"
                    checked={allSelected}
                    onChange={toggleSelectAll}
                    aria-label="Select all users"
                  />
                </th>
                <th>Name</th>
                <th>Position</th>
                <th>Phone Number</th>
                <th>Email</th>
                <th>Role</th>
              </tr>
            </thead>
            <tbody>
              {users.map((user) => {
                const isSelf = user.email === signedInEmail;
                const isSelected = selectedEmails.includes(user.email);
                return (
                  <tr
                    key={user.email}
                    className={
                      isSelected
                        ? "row-selectable row-selected"
                        : "row-selectable"
                    }
                    onClick={() => toggleSelected(user.email)}
                  >
                    <td className="col-check">
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => toggleSelected(user.email)}
                        onClick={(event) => event.stopPropagation()}
                        aria-label={`Select ${user.email}`}
                      />
                    </td>
                    <td>{user.name || "—"}</td>
                    <td>{user.position || "—"}</td>
                    <td>{user.phone || "—"}</td>
                    <td>
                      {user.email}
                      {isSelf ? (
                        <span className="lookup-meta"> (you)</span>
                      ) : null}
                    </td>
                    <td>{ROLE_LABELS[user.role]}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {editor && mounted
        ? createPortal(
            <div className="dialog-backdrop" role="presentation">
              <div
                className="dialog-panel dialog-panel--wide"
                role="dialog"
                aria-modal="true"
                aria-labelledby="user-editor-title"
              >
                <div className="dialog-header">
                  <h3 id="user-editor-title" className="dialog-title">
                    {editor.mode === "add" ? "Add User" : "Edit User"}
                  </h3>
                  <button
                    type="button"
                    className="dialog-close"
                    aria-label="Close"
                    onClick={closeEditor}
                    disabled={saving}
                  >
                    ×
                  </button>
                </div>

                {error ? <p className="form-message error">{error}</p> : null}

                <form
                  className="client-form client-form--dialog"
                  onSubmit={(event) => void submitEditor(event)}
                >
                  <div className="client-form-grid">
                    <label className="field">
                      <span className="field-label">Name</span>
                      <input
                        className="field-input"
                        value={editor.name}
                        placeholder="Full name"
                        disabled={saving}
                        onChange={(event) =>
                          setEditor({ ...editor, name: event.target.value })
                        }
                        autoFocus
                      />
                    </label>

                    <label className="field">
                      <span className="field-label">Position</span>
                      <input
                        className="field-input"
                        value={editor.position}
                        placeholder="Principal Operations, Consultant"
                        disabled={saving}
                        onChange={(event) =>
                          setEditor({
                            ...editor,
                            position: event.target.value,
                          })
                        }
                      />
                    </label>

                    <label className="field">
                      <span className="field-label">Phone Number</span>
                      <input
                        className="field-input"
                        value={editor.phone}
                        placeholder="(559) 333 2167"
                        disabled={saving}
                        onChange={(event) =>
                          setEditor({ ...editor, phone: event.target.value })
                        }
                      />
                    </label>

                    <label className="field">
                      <span className="field-label">
                        Email <span className="field-required">*</span>
                      </span>
                      <input
                        className={
                          editor.mode === "edit"
                            ? "field-input field-input--readonly"
                            : "field-input"
                        }
                        type="email"
                        value={editor.email}
                        placeholder="name@fireflyaccountancy.com"
                        disabled={saving}
                        readOnly={editor.mode === "edit"}
                        tabIndex={editor.mode === "edit" ? -1 : undefined}
                        required
                        onChange={(event) =>
                          setEditor({ ...editor, email: event.target.value })
                        }
                      />
                    </label>

                    <label className="field">
                      <span className="field-label">Role</span>
                      <select
                        className="field-input field-select"
                        value={editor.role}
                        disabled={
                          saving ||
                          (editor.mode === "edit" &&
                            editor.role === "admin" &&
                            adminCount <= 1)
                        }
                        onChange={(event) =>
                          setEditor({
                            ...editor,
                            role: event.target.value as PortalRole,
                          })
                        }
                      >
                        {PORTAL_ROLES.map((role) => (
                          <option key={role} value={role}>
                            {ROLE_LABELS[role]}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>

                  <div className="dialog-actions">
                    <button
                      type="button"
                      className="button secondary"
                      onClick={closeEditor}
                      disabled={saving}
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      className="button primary"
                      disabled={saving}
                    >
                      {saving ? "Saving…" : "Save"}
                    </button>
                  </div>
                </form>
              </div>
            </div>,
            document.body,
          )
        : null}
    </section>
  );
}
