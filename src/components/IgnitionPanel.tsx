"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { useAccess } from "@/auth/accessContext";
import {
  IgnitionProposalViewer,
  type IgnitionProposalViewerTarget,
} from "@/components/IgnitionProposalViewer";
import {
  endingSoonEmailSendersLabel,
  isEndingSoonEmailSender,
  type EndingSoonEmailStatusLabel,
} from "@/lib/ignition-ending-soon-email-ui";
import { portalFetch } from "@/lib/portal-api";

type IgnitionClientRow = {
  id: string;
  name: string;
  email: string;
  status: string;
};

type EndingSoonRow = {
  slug: string;
  clientSlug: string;
  clientName: string;
  name: string;
  referenceNumber: string;
  state: string;
  startDate: string;
  minimumContractLength: number | null;
  endsOn: string;
  viewUrl: string;
};

type InvoiceSummaryRow = {
  clientSlug: string;
  clientName: string;
  lastPaidOn: string;
  lastPaidAmountCents: number | null;
  lastPaidCurrency: string;
  lastPaidInvoiceNumber: string;
  previousDuesOn: string;
  previousDuesAmountCents: number | null;
  previousDuesCurrency: string;
  previousDuesInvoiceNumber: string;
  previousDuesCount: number;
  nextUpcomingOn: string;
  nextUpcomingAmountCents: number | null;
  nextUpcomingCurrency: string;
  nextUpcomingInvoiceNumber: string;
};

type NewClientFromAcceptedProposal = {
  profileId: string;
  uniqueId: string;
  legalName: string;
  ignitionClientSlug: string;
  proposalSlug: string;
  proposalReference: string;
};

type IgnitionStatusResponse = {
  connected?: boolean;
  configured?: boolean;
  expiresAt?: string | null;
  syncedAt?: string | null;
  endingSoonSyncedAt?: string | null;
  invoicesSyncedAt?: string | null;
  clients?: IgnitionClientRow[];
  endingSoon?: EndingSoonRow[];
  invoiceSummaries?: InvoiceSummaryRow[];
  newClientsFromAcceptedProposals?: NewClientFromAcceptedProposal[];
  error?: string;
};

type IgnitionInnerTab = "overview" | "endingSoon" | "invoices";

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

function formatDateOnly(value?: string | null) {
  if (!value) {
    return "—";
  }
  const raw = value.trim();
  const date = new Date(raw.includes("T") ? raw : `${raw}T00:00:00`);
  if (Number.isNaN(date.getTime())) {
    return raw;
  }
  return date.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

function formatInvoiceAmount(
  amountCents: number | null | undefined,
  currency = "USD",
): string {
  if (amountCents == null || !Number.isFinite(amountCents)) {
    return "—";
  }
  const amount = amountCents / 100;
  try {
    return new Intl.NumberFormat(undefined, {
      style: "currency",
      currency: currency || "USD",
    }).format(amount);
  } catch {
    return `${currency || "USD"} ${amount.toFixed(2)}`;
  }
}

function emailStatusClass(label: EndingSoonEmailStatusLabel | undefined) {
  if (label === "Sent") {
    return "pakistan-holiday-email-status pakistan-holiday-email-status--sent";
  }
  if (label === "Scheduled") {
    return "pakistan-holiday-email-status pakistan-holiday-email-status--scheduled";
  }
  return "pakistan-holiday-email-status";
}

export function IgnitionPanel() {
  const { canManageUsers, email: signedInEmail } = useAccess();
  const [innerTab, setInnerTab] = useState<IgnitionInnerTab>("overview");
  const [connected, setConnected] = useState(false);
  const [configured, setConfigured] = useState(false);
  const [expiresAt, setExpiresAt] = useState<string | null>(null);
  const [syncedAt, setSyncedAt] = useState<string | null>(null);
  const [endingSoonSyncedAt, setEndingSoonSyncedAt] = useState<string | null>(
    null,
  );
  const [invoicesSyncedAt, setInvoicesSyncedAt] = useState<string | null>(null);
  const [clients, setClients] = useState<IgnitionClientRow[]>([]);
  const [endingSoon, setEndingSoon] = useState<EndingSoonRow[]>([]);
  const [invoiceSummaries, setInvoiceSummaries] = useState<InvoiceSummaryRow[]>(
    [],
  );
  const [teamEmailStatusBySlug, setTeamEmailStatusBySlug] = useState<
    Record<string, EndingSoonEmailStatusLabel>
  >({});
  const [clientEmailStatusBySlug, setClientEmailStatusBySlug] = useState<
    Record<string, EndingSoonEmailStatusLabel>
  >({});
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState<
    "connect" | "disconnect" | "sync" | "webhook" | null
  >(null);
  const [schedulingEmails, setSchedulingEmails] = useState(false);
  const [webhookRegistered, setWebhookRegistered] = useState(false);
  const [webhookHasSecret, setWebhookHasSecret] = useState(false);
  const [webhookReceiveUrl, setWebhookReceiveUrl] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [proposalViewer, setProposalViewer] =
    useState<IgnitionProposalViewerTarget | null>(null);
  const canScheduleEndingSoonEmails = isEndingSoonEmailSender(signedInEmail);
  const endingSoonEnsureStarted = useRef(false);

  function openEndingSoonProposal(
    proposal: EndingSoonRow,
    mode: "view" | "download",
  ) {
    const slug = proposal.slug.trim();
    if (!slug) {
      return;
    }
    setProposalViewer({
      slug,
      reference:
        proposal.referenceNumber.trim() ||
        proposal.name.trim() ||
        slug,
      clientLabel: proposal.clientName.trim() || "Client",
      mode,
    });
  }

  const applyStatus = useCallback((data: IgnitionStatusResponse) => {
    setConnected(Boolean(data.connected));
    setConfigured(Boolean(data.configured));
    setExpiresAt(data.expiresAt ?? null);
    setSyncedAt(data.syncedAt ?? null);
    setEndingSoonSyncedAt(data.endingSoonSyncedAt ?? null);
    setInvoicesSyncedAt(data.invoicesSyncedAt ?? null);
    setClients(Array.isArray(data.clients) ? data.clients : []);
    setEndingSoon(Array.isArray(data.endingSoon) ? data.endingSoon : []);
    setInvoiceSummaries(
      Array.isArray(data.invoiceSummaries) ? data.invoiceSummaries : [],
    );
  }, []);

  const loadWebhookStatus = useCallback(async () => {
    try {
      const response = await portalFetch("/api/ignition/webhooks/manage", {
        cache: "no-store",
      });
      const data = (await response.json()) as {
        registered?: boolean;
        hasSecret?: boolean;
        receiveUrl?: string;
      };
      if (!response.ok) {
        return;
      }
      setWebhookRegistered(Boolean(data.registered));
      setWebhookHasSecret(Boolean(data.hasSecret));
      setWebhookReceiveUrl(String(data.receiveUrl ?? "").trim());
    } catch {
      // Optional until webhooks are enabled on the server.
    }
  }, []);

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
      applyStatus(data);
      await loadWebhookStatus();
    } catch (loadError) {
      setError(
        loadError instanceof Error
          ? loadError.message
          : "Failed to load Ignition status.",
      );
    } finally {
      setLoading(false);
    }
  }, [applyStatus, loadWebhookStatus]);

  const loadEndingSoonEmailStatus = useCallback(async () => {
    try {
      const response = await portalFetch("/api/ignition/ending-soon-emails", {
        cache: "no-store",
      });
      const data = (await response.json()) as {
        teamStatusBySlug?: Record<string, EndingSoonEmailStatusLabel>;
        clientStatusBySlug?: Record<string, EndingSoonEmailStatusLabel>;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to load Ending Soon email status.");
      }
      setTeamEmailStatusBySlug(
        data.teamStatusBySlug && typeof data.teamStatusBySlug === "object"
          ? data.teamStatusBySlug
          : {},
      );
      setClientEmailStatusBySlug(
        data.clientStatusBySlug && typeof data.clientStatusBySlug === "object"
          ? data.clientStatusBySlug
          : {},
      );
    } catch {
      // Status columns stay empty; ignition status still loads.
    }
  }, []);

  /** Status/logic only — never sends mail. */
  const ensureEndingSoonEmails = useCallback(async () => {
    if (!canScheduleEndingSoonEmails || schedulingEmails) {
      await loadEndingSoonEmailStatus();
      return;
    }
    setSchedulingEmails(true);
    setError(null);
    try {
      const response = await portalFetch("/api/ignition/ending-soon-emails", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "ensure" }),
      });
      const data = (await response.json()) as {
        teamStatusBySlug?: Record<string, EndingSoonEmailStatusLabel>;
        clientStatusBySlug?: Record<string, EndingSoonEmailStatusLabel>;
        message?: string;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(
          data.error || "Failed to update Ending Soon email status.",
        );
      }
      setTeamEmailStatusBySlug(
        data.teamStatusBySlug && typeof data.teamStatusBySlug === "object"
          ? data.teamStatusBySlug
          : {},
      );
      setClientEmailStatusBySlug(
        data.clientStatusBySlug && typeof data.clientStatusBySlug === "object"
          ? data.clientStatusBySlug
          : {},
      );
      setNotice(
        data.message ||
          "Team and Client emails marked Scheduled. Nothing was sent.",
      );
    } catch (ensureError) {
      setError(
        ensureError instanceof Error
          ? ensureError.message
          : "Failed to update Ending Soon email status.",
      );
      await loadEndingSoonEmailStatus();
    } finally {
      setSchedulingEmails(false);
    }
  }, [
    canScheduleEndingSoonEmails,
    loadEndingSoonEmailStatus,
    schedulingEmails,
  ]);

  useEffect(() => {
    if (!canManageUsers) {
      setLoading(false);
      return;
    }
    void loadStatus();
  }, [canManageUsers, loadStatus]);

  useEffect(() => {
    endingSoonEnsureStarted.current = false;
  }, [endingSoonSyncedAt]);

  useEffect(() => {
    if (
      !canManageUsers ||
      innerTab !== "endingSoon" ||
      endingSoon.length === 0 ||
      !canScheduleEndingSoonEmails
    ) {
      return;
    }
    if (endingSoonEnsureStarted.current) {
      return;
    }
    endingSoonEnsureStarted.current = true;
    void ensureEndingSoonEmails();
  }, [
    canManageUsers,
    canScheduleEndingSoonEmails,
    endingSoon.length,
    ensureEndingSoonEmails,
    innerTab,
  ]);

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
      setEndingSoon([]);
      setInvoiceSummaries([]);
      setInvoicesSyncedAt(null);
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

  async function handleEnableWebhooks() {
    setBusy("webhook");
    setError(null);
    setNotice(null);
    try {
      const response = await portalFetch("/api/ignition/webhooks/manage", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "ensure" }),
      });
      const data = (await response.json()) as {
        message?: string;
        registered?: boolean;
        hasSecret?: boolean;
        receiveUrl?: string;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to register Ignition webhook.");
      }
      setWebhookRegistered(Boolean(data.registered));
      setWebhookHasSecret(Boolean(data.hasSecret));
      setWebhookReceiveUrl(String(data.receiveUrl ?? "").trim());
      setNotice(
        data.message ||
          "Ignition webhook enabled. Accepted proposals will notify the team automatically.",
      );
    } catch (webhookError) {
      setError(
        webhookError instanceof Error
          ? webhookError.message
          : "Failed to register Ignition webhook.",
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
      setEndingSoonSyncedAt(data.endingSoonSyncedAt ?? null);
      setInvoicesSyncedAt(data.invoicesSyncedAt ?? null);
      setClients(Array.isArray(data.clients) ? data.clients : []);
      setEndingSoon(Array.isArray(data.endingSoon) ? data.endingSoon : []);
      setInvoiceSummaries(
        Array.isArray(data.invoiceSummaries) ? data.invoiceSummaries : [],
      );
      const clientCount = Array.isArray(data.clients) ? data.clients.length : 0;
      const endingCount = Array.isArray(data.endingSoon)
        ? data.endingSoon.length
        : 0;
      const invoiceClientCount = Array.isArray(data.invoiceSummaries)
        ? data.invoiceSummaries.length
        : 0;
      const newClients = Array.isArray(data.newClientsFromAcceptedProposals)
        ? data.newClientsFromAcceptedProposals
        : [];
      const parts = [
        `${clientCount} client(s)`,
        `${endingCount} ending-soon proposal(s)`,
        `invoices for ${invoiceClientCount} client(s)`,
      ];
      if (newClients.length > 0) {
        parts.push(
          `${newClients.length} new Client Info profile(s) from accepted proposals`,
        );
        window.dispatchEvent(new Event("dashboard-notifications-refresh"));
      }
      setNotice(`Synced ${parts.join(", ")}.`);
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
            Connect the Ignition Core API, then sync clients and accepted
            proposal PDFs into Client Profile.
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
                {busy === "sync" ? "Syncing…" : "Sync clients & proposals"}
              </button>
              <button
                type="button"
                className="button secondary-solid"
                disabled={busy !== null}
                onClick={() => void handleEnableWebhooks()}
                title={
                  webhookReceiveUrl
                    ? `Receives proposal.accepted at ${webhookReceiveUrl}`
                    : "Register Ignition webhook on this server"
                }
              >
                {busy === "webhook"
                  ? "Enabling…"
                  : webhookRegistered && webhookHasSecret
                    ? "Refresh auto notifications"
                    : "Enable auto notifications"}
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
        <div>
          <dt>Auto notifications</dt>
          <dd>
            {webhookRegistered && webhookHasSecret
              ? "Enabled (proposal.accepted webhook)"
              : connected
                ? "Not enabled — click Enable auto notifications on the server"
                : "—"}
          </dd>
        </div>
      </dl>

      {!configured ? (
        <p className="lookup-hint">
          Add <code>IGNITION_CLIENT_ID</code> and{" "}
          <code>IGNITION_CLIENT_SECRET</code> to <code>.env.local</code> (or
          server <code>.env</code>), then restart the app.
        </p>
      ) : (
        <p className="lookup-hint">
          Sync attaches each client&apos;s accepted proposal to Client Profile
          (Active Proposal column). Ending Soon lists proposals that expire in
          the next 3 months. Invoices shows last paid, previous dues, and next
          upcoming bill per client from Ignition. On the live server, use{" "}
          <strong>Enable auto notifications</strong> so Ignition pushes{" "}
          <code>proposal.accepted</code> to{" "}
          <code>
            {webhookReceiveUrl ||
              "https://clientportal.fireflyaccountancy.com/api/ignition/webhooks"}
          </code>{" "}
          — new clients are added to Client Info and the team sees a dashboard
          banner.
        </p>
      )}

      <div
        className="lookup-section-tabs"
        role="tablist"
        aria-label="Ignition sections"
      >
        <button
          type="button"
          role="tab"
          aria-selected={innerTab === "overview"}
          className={
            innerTab === "overview"
              ? "lookup-section-tab active"
              : "lookup-section-tab"
          }
          onClick={() => setInnerTab("overview")}
        >
          Overview
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={innerTab === "endingSoon"}
          className={
            innerTab === "endingSoon"
              ? "lookup-section-tab active"
              : "lookup-section-tab"
          }
          onClick={() => setInnerTab("endingSoon")}
        >
          Ending Soon
          {endingSoon.length > 0 ? ` (${endingSoon.length})` : ""}
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={innerTab === "invoices"}
          className={
            innerTab === "invoices"
              ? "lookup-section-tab active"
              : "lookup-section-tab"
          }
          onClick={() => setInnerTab("invoices")}
        >
          Invoices
          {invoiceSummaries.length > 0 ? ` (${invoiceSummaries.length})` : ""}
        </button>
      </div>

      {innerTab === "overview" ? (
        <div className="table-wrap">
          {loading ? (
            <p className="table-empty">Loading Ignition status…</p>
          ) : clients.length === 0 ? (
            <p className="table-empty">
              {connected
                ? "No clients synced yet. Click Sync clients & proposals."
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
      ) : innerTab === "endingSoon" ? (
        <div className="table-wrap">
          {loading ? (
            <p className="table-empty">Loading ending-soon proposals…</p>
          ) : endingSoon.length === 0 ? (
            <p className="table-empty">
              {connected
                ? "No proposals ending in the next 3 months. Run Sync to refresh."
                : "Connect to Ignition, then sync to load ending-soon proposals."}
            </p>
          ) : (
            <>
              <p className="lookup-hint">
                Proposals ending in the next 3 months
                {endingSoonSyncedAt
                  ? ` · last refreshed ${formatWhen(endingSoonSyncedAt)}`
                  : ""}
                . The planned send time is calculated and stored, but nothing is
                queued in Outlook and no mail goes out. Team: 2 months before @{" "}
                <strong>9:00 PM Pakistan time</strong> (2027+ → next working
                day). Client: 1 month before, same time. Weekend/holiday → next
                working day. From: {endingSoonEmailSendersLabel()}.
                {schedulingEmails ? " · updating status…" : ""}
              </p>
              <table className="projects-table">
                <thead>
                  <tr>
                    <th>Client</th>
                    <th>Reference</th>
                    <th>Term (mo)</th>
                    <th>Ends on</th>
                    <th>Team (Email)</th>
                    <th>Client (Email)</th>
                  </tr>
                </thead>
                <tbody>
                  {endingSoon.map((proposal) => {
                    const teamStatus =
                      teamEmailStatusBySlug[proposal.slug] ?? "—";
                    const clientStatus =
                      clientEmailStatusBySlug[proposal.slug] ?? "—";
                    return (
                      <tr key={proposal.slug}>
                        <td>{proposal.clientName || "—"}</td>
                        <td>
                          {proposal.slug &&
                          (proposal.referenceNumber || proposal.name) ? (
                            <button
                              type="button"
                              className="table-link-button"
                              title={proposal.name || "View proposal"}
                              onClick={() =>
                                openEndingSoonProposal(proposal, "view")
                              }
                            >
                              {proposal.referenceNumber ? (
                                <code>{proposal.referenceNumber}</code>
                              ) : (
                                proposal.name || "Proposal"
                              )}
                            </button>
                          ) : proposal.referenceNumber ? (
                            <code>{proposal.referenceNumber}</code>
                          ) : (
                            "—"
                          )}
                        </td>
                        <td>
                          {proposal.minimumContractLength != null
                            ? proposal.minimumContractLength
                            : "—"}
                        </td>
                        <td>{formatDateOnly(proposal.endsOn)}</td>
                        <td>
                          <span className={emailStatusClass(teamStatus)}>
                            {teamStatus}
                          </span>
                        </td>
                        <td>
                          <span className={emailStatusClass(clientStatus)}>
                            {clientStatus}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </>
          )}
        </div>
      ) : (
        <div className="table-wrap">
          {loading ? (
            <p className="table-empty">Loading invoices…</p>
          ) : invoiceSummaries.length === 0 ? (
            <p className="table-empty">
              {connected
                ? "No invoices synced yet. Click Sync clients & proposals."
                : "Connect to Ignition, then sync to load invoices."}
            </p>
          ) : (
            <>
              <p className="lookup-hint">
                Per client from Ignition invoices
                {invoicesSyncedAt
                  ? ` · last refreshed ${formatWhen(invoicesSyncedAt)}`
                  : ""}
                . Last paid = most recent paid invoice. Previous dues = unpaid
                invoices past due (amount is the total overdue). Next upcoming
                bill = soonest unpaid invoice due today or later.
              </p>
              <table className="projects-table">
                <thead>
                  <tr>
                    <th>Client</th>
                    <th>Last paid</th>
                    <th>Last amount</th>
                    <th>Previous dues</th>
                    <th>Previous amount</th>
                    <th>Next upcoming</th>
                    <th>Upcoming amount</th>
                  </tr>
                </thead>
                <tbody>
                  {invoiceSummaries.map((row) => (
                    <tr key={row.clientSlug}>
                      <td>{row.clientName || "—"}</td>
                      <td>
                        {row.lastPaidOn ? (
                          <>
                            {formatDateOnly(row.lastPaidOn)}
                            {row.lastPaidInvoiceNumber ? (
                              <>
                                {" "}
                                <code>{row.lastPaidInvoiceNumber}</code>
                              </>
                            ) : null}
                          </>
                        ) : (
                          "—"
                        )}
                      </td>
                      <td>
                        {formatInvoiceAmount(
                          row.lastPaidAmountCents,
                          row.lastPaidCurrency,
                        )}
                      </td>
                      <td>
                        {row.previousDuesOn ? (
                          <>
                            {formatDateOnly(row.previousDuesOn)}
                            {row.previousDuesInvoiceNumber ? (
                              <>
                                {" "}
                                <code>{row.previousDuesInvoiceNumber}</code>
                              </>
                            ) : null}
                          </>
                        ) : (
                          "—"
                        )}
                      </td>
                      <td>
                        {formatInvoiceAmount(
                          row.previousDuesAmountCents,
                          row.previousDuesCurrency,
                        )}
                      </td>
                      <td>
                        {row.nextUpcomingOn ? (
                          <>
                            {formatDateOnly(row.nextUpcomingOn)}
                            {row.nextUpcomingInvoiceNumber ? (
                              <>
                                {" "}
                                <code>{row.nextUpcomingInvoiceNumber}</code>
                              </>
                            ) : null}
                          </>
                        ) : (
                          "—"
                        )}
                      </td>
                      <td>
                        {formatInvoiceAmount(
                          row.nextUpcomingAmountCents,
                          row.nextUpcomingCurrency,
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          )}
        </div>
      )}

      <IgnitionProposalViewer
        target={proposalViewer}
        onClose={() => setProposalViewer(null)}
      />
    </section>
  );
}
