"use client";

import { useCallback, useEffect, useState } from "react";
import { useAccess } from "@/auth/accessContext";
import { portalFetch } from "@/lib/portal-api";

type IgnitionClientRow = {
  id: string;
  name: string;
  email: string;
  status: string;
};

type IgnitionStatusResponse = {
  connected?: boolean;
  configured?: boolean;
  expiresAt?: string | null;
  syncedAt?: string | null;
  clients?: IgnitionClientRow[];
  error?: string;
};

function formatWhen(value?: string | null) {
  if (!value) {
    return "—";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleString();
}

export function IgnitionPanel() {
  const { canManageUsers } = useAccess();
  const [connected, setConnected] = useState(false);
  const [configured, setConfigured] = useState(false);
  const [expiresAt, setExpiresAt] = useState<string | null>(null);
  const [syncedAt, setSyncedAt] = useState<string | null>(null);
  const [clients, setClients] = useState<IgnitionClientRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState<"connect" | "disconnect" | "sync" | null>(
    null,
  );
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  const loadStatus = useCallback(async () => {
    setError(null);
    try {
      const response = await portalFetch("/api/ignition/status", {
        cache: "no-store",
      });
      const data = (await response.json()) as IgnitionStatusResponse;
      if (!response.ok) {
        throw new Error(data.error || "Failed to load Ignition status.");
      }
      setConnected(Boolean(data.connected));
      setConfigured(Boolean(data.configured));
      setExpiresAt(data.expiresAt ?? null);
      setSyncedAt(data.syncedAt ?? null);
      setClients(Array.isArray(data.clients) ? data.clients : []);
    } catch (loadError) {
      setError(
        loadError instanceof Error
          ? loadError.message
          : "Failed to load Ignition status.",
      );
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!canManageUsers) {
      setLoading(false);
      return;
    }
    void loadStatus();
  }, [canManageUsers, loadStatus]);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const result = params.get("ignition");
    if (!result) {
      return;
    }

    if (result === "connected") {
      setNotice("Connected to Ignition.");
    } else if (result === "error") {
      setError(params.get("message") || "Ignition connection failed.");
    }

    params.delete("ignition");
    params.delete("message");
    const next = params.toString();
    window.history.replaceState({}, "", next ? `/?${next}` : "/");
  }, []);

  async function handleConnect() {
    setBusy("connect");
    setError(null);
    setNotice(null);
    try {
      const response = await portalFetch("/api/ignition/connect", {
        cache: "no-store",
      });
      const data = (await response.json()) as {
        authorizeUrl?: string;
        error?: string;
      };
      if (!response.ok || !data.authorizeUrl) {
        throw new Error(data.error || "Failed to start Ignition connect.");
      }
      window.location.href = data.authorizeUrl;
    } catch (connectError) {
      setBusy(null);
      setError(
        connectError instanceof Error
          ? connectError.message
          : "Failed to start Ignition connect.",
      );
    }
  }

  async function handleDisconnect() {
    if (!window.confirm("Disconnect Ignition from this portal?")) {
      return;
    }
    setBusy("disconnect");
    setError(null);
    setNotice(null);
    try {
      const response = await portalFetch("/api/ignition/disconnect", {
        method: "POST",
      });
      const data = (await response.json()) as { error?: string };
      if (!response.ok) {
        throw new Error(data.error || "Failed to disconnect Ignition.");
      }
      setConnected(false);
      setExpiresAt(null);
      setNotice("Ignition disconnected.");
    } catch (disconnectError) {
      setError(
        disconnectError instanceof Error
          ? disconnectError.message
          : "Failed to disconnect Ignition.",
      );
    } finally {
      setBusy(null);
    }
  }

  async function handleSync() {
    setBusy("sync");
    setError(null);
    setNotice(null);
    try {
      const response = await portalFetch("/api/ignition/sync/clients", {
        method: "POST",
      });
      const data = (await response.json()) as IgnitionStatusResponse;
      if (!response.ok) {
        throw new Error(data.error || "Failed to sync Ignition clients.");
      }
      setSyncedAt(data.syncedAt ?? null);
      setClients(Array.isArray(data.clients) ? data.clients : []);
      setNotice(
        `Synced ${Array.isArray(data.clients) ? data.clients.length : 0} Ignition client(s).`,
      );
    } catch (syncError) {
      setError(
        syncError instanceof Error
          ? syncError.message
          : "Failed to sync Ignition clients.",
      );
    } finally {
      setBusy(null);
    }
  }

  if (!canManageUsers) {
    return (
      <section className="content-panel">
        <p className="lookup-hint">Admin access is required for Ignition.</p>
      </section>
    );
  }

  return (
    <section className="content-panel content-panel--actions">
      <div className="panel-title-row users-header">
        <div className="users-header-copy">
          <h2 className="section-title">Ignition</h2>
          <p className="lookup-hint">
            Connect the Ignition Core API, then sync clients into this portal.
          </p>
        </div>
        <div className="panel-title-actions">
          {!connected ? (
            <button
              type="button"
              className="button primary"
              disabled={busy !== null || !configured}
              onClick={() => void handleConnect()}
            >
              {busy === "connect" ? "Connecting…" : "Connect to Ignition"}
            </button>
          ) : (
            <>
              <button
                type="button"
                className="button secondary-solid"
                disabled={busy !== null}
                onClick={() => void handleSync()}
              >
                {busy === "sync" ? "Syncing…" : "Sync clients"}
              </button>
              <button
                type="button"
                className="button secondary"
                disabled={busy !== null}
                onClick={() => void handleDisconnect()}
              >
                {busy === "disconnect" ? "Disconnecting…" : "Disconnect"}
              </button>
            </>
          )}
        </div>
      </div>

      {notice ? (
        <div className="inline-notice inline-notice--success" role="status">
          <p className="inline-notice__text">{notice}</p>
          <button
            type="button"
            className="inline-notice__close"
            aria-label="Dismiss"
            onClick={() => setNotice(null)}
          >
            ×
          </button>
        </div>
      ) : null}

      {error ? <p className="form-message error">{error}</p> : null}

      <dl className="ignition-status">
        <div>
          <dt>Configuration</dt>
          <dd>{configured ? "Credentials set" : "Missing env credentials"}</dd>
        </div>
        <div>
          <dt>Connection</dt>
          <dd>{connected ? "Connected" : "Not connected"}</dd>
        </div>
        <div>
          <dt>Token expires</dt>
          <dd>{formatWhen(expiresAt)}</dd>
        </div>
        <div>
          <dt>Last client sync</dt>
          <dd>{formatWhen(syncedAt)}</dd>
        </div>
      </dl>

      {!configured ? (
        <p className="lookup-hint">
          Add <code>IGNITION_CLIENT_ID</code> and{" "}
          <code>IGNITION_CLIENT_SECRET</code> to <code>.env.local</code> (or
          server <code>.env</code>), then restart the app.
        </p>
      ) : null}

      <div className="table-wrap">
        {loading ? (
          <p className="table-empty">Loading Ignition status…</p>
        ) : clients.length === 0 ? (
          <p className="table-empty">
            {connected
              ? "No clients synced yet. Click Sync clients."
              : "Connect to Ignition to sync clients."}
          </p>
        ) : (
          <table className="projects-table">
            <thead>
              <tr>
                <th>Name</th>
                <th>Email</th>
                <th>Ignition ID</th>
                <th>Status</th>
                <th>Last synced</th>
              </tr>
            </thead>
            <tbody>
              {clients.map((client) => (
                <tr key={client.id}>
                  <td>{client.name || "—"}</td>
                  <td>{client.email || "—"}</td>
                  <td>
                    <code>{client.id}</code>
                  </td>
                  <td>{client.status || "—"}</td>
                  <td>{formatWhen(syncedAt)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </section>
  );
}
