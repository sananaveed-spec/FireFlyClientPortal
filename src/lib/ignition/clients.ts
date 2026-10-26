import { copyFile, mkdir, readFile, unlink, writeFile } from "fs/promises";
import path from "path";
import {
  getDataDir,
  readClientProfiles,
  writeClientProfiles,
} from "@/lib/data-store";
import { getIgnitionConfig } from "@/lib/ignition/config";
import { appendDashboardNotifications } from "@/lib/dashboard-notifications";
import { appendHistory } from "@/lib/data-store";
import { syncIgnitionInvoices } from "@/lib/ignition/invoices";
import {
  applyActiveProposalsToProfiles,
  createProfilesForAcceptedIgnitionClients,
  syncIgnitionEndingSoon,
  syncIgnitionProposals,
  type IgnitionAutoCreatedClientProfile,
} from "@/lib/ignition/proposals";
import { getValidAccessToken } from "@/lib/ignition/tokens";

export type IgnitionClientsSyncResult = IgnitionClientsStore & {
  newClientsFromAcceptedProposals: IgnitionAutoCreatedClientProfile[];
};

export type IgnitionClientRecord = {
  id: string;
  name: string;
  email: string;
  status: string;
  raw: Record<string, unknown>;
};

export type IgnitionClientsStore = {
  syncedAt: string;
  clients: IgnitionClientRecord[];
};

function clientsPath() {
  return path.join(getDataDir(), "ignition-clients.json");
}

async function ensureDataDir() {
  await mkdir(getDataDir(), { recursive: true });
}

async function writeJsonFile(filePath: string, value: unknown) {
  const tempPath = `${filePath}.${process.pid}.tmp`;
  const payload = `${JSON.stringify(value, null, 2)}\n`;
  await writeFile(tempPath, payload, "utf8");
  await copyFile(tempPath, filePath);
  await unlink(tempPath).catch(() => undefined);
}

function isMissingFile(error: unknown) {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    (error as { code?: string }).code === "ENOENT"
  );
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

function pickString(record: Record<string, unknown>, keys: string[]): string {
  for (const key of keys) {
    const value = record[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
    if (typeof value === "number" && Number.isFinite(value)) {
      return String(value);
    }
  }
  return "";
}

function pickEmail(record: Record<string, unknown>): string {
  const direct = pickString(record, [
    "email",
    "primary_email",
    "primaryEmail",
    "email_address",
    "emailAddress",
  ]);
  if (direct) {
    return direct;
  }

  const contacts = record.contacts;
  if (Array.isArray(contacts)) {
    for (const contact of contacts) {
      const row = asRecord(contact);
      if (!row) {
        continue;
      }
      const email = pickString(row, ["email", "email_address", "emailAddress"]);
      if (email) {
        return email;
      }
    }
  }

  return "";
}

export function normalizeIgnitionClient(
  value: unknown,
): IgnitionClientRecord | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }

  // Core API identifies clients by slug (e.g. cli_...).
  const id = pickString(record, [
    "slug",
    "id",
    "uuid",
    "client_id",
    "clientId",
  ]);
  if (!id) {
    return null;
  }

  const name =
    pickString(record, [
      "name",
      "client_name",
      "clientName",
      "legal_name",
      "legalName",
      "display_name",
      "displayName",
      "company_name",
      "companyName",
    ]) || id;

  return {
    id,
    name,
    email: pickEmail(record),
    status: pickString(record, ["status", "state", "lifecycle_state"]),
    raw: record,
  };
}

function normalizeIgnitionClientError(
  message: string,
  status: number,
  code = "",
): string {
  const lower = `${message} ${code}`.toLowerCase();
  if (status === 403 || lower.includes("core_api_disabled")) {
    return [
      "Ignition Core API is disabled for this practice.",
      "An Admin must turn it on under Settings → Labs → Core API,",
      "then Disconnect and Connect again in this portal.",
    ].join(" ");
  }
  if (
    lower.includes("valid email and password") ||
    (status === 401 &&
      (lower.includes("unauthorized") || lower.includes("password")))
  ) {
    return [
      "Ignition rejected the client sync request (HTTP " + status + ").",
      "Disconnect and Connect again so the portal can obtain a Core API token with PKCE.",
      "Confirm Core API is on under Labs and the Developer hub app is Core API (not Reporting).",
    ].join(" ");
  }
  return message;
}

function extractClientRows(payload: unknown): unknown[] {
  if (Array.isArray(payload)) {
    return payload;
  }
  const record = asRecord(payload);
  if (!record) {
    return [];
  }

  for (const key of ["data", "clients", "items", "results", "records"]) {
    const value = record[key];
    if (Array.isArray(value)) {
      return value;
    }
  }

  return [];
}

function listHasMore(payload: unknown): boolean {
  const record = asRecord(payload);
  return Boolean(record && record.has_more === true);
}

async function fetchClientsPage(
  url: string,
  accessToken: string,
): Promise<unknown> {
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${accessToken}`,
    },
    cache: "no-store",
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const record = asRecord(payload);
    const errorObj = record ? asRecord(record.error) : null;
    const code =
      (errorObj && pickString(errorObj, ["code", "error", "type"])) ||
      (record && pickString(record, ["code", "error_code", "type"])) ||
      "";
    const raw =
      (record &&
        (pickString(record, ["error_description", "message", "error"]) ||
          (errorObj &&
            pickString(errorObj, ["message", "error_description", "code"])) ||
          JSON.stringify(record))) ||
      `Ignition clients request failed (${response.status}).`;
    throw new Error(normalizeIgnitionClientError(raw, response.status, code));
  }

  return payload;
}

export async function fetchAllIgnitionClients(): Promise<IgnitionClientRecord[]> {
  const { accessToken } = await getValidAccessToken();
  const { apiBase } = getIgnitionConfig();

  const clients: IgnitionClientRecord[] = [];
  const seenIds = new Set<string>();
  let startingAfter = "";
  let pages = 0;

  while (pages < 50) {
    pages += 1;
    const url = new URL(`${apiBase}/clients`);
    url.searchParams.set("limit", "100");
    if (startingAfter) {
      url.searchParams.set("starting_after", startingAfter);
    }

    const payload = await fetchClientsPage(url.toString(), accessToken);
    const rows = extractClientRows(payload);
    if (rows.length === 0) {
      break;
    }

    let lastSlug = "";
    for (const row of rows) {
      const client = normalizeIgnitionClient(row);
      if (!client || seenIds.has(client.id)) {
        continue;
      }
      seenIds.add(client.id);
      clients.push(client);
      lastSlug = client.id;
    }

    if (!listHasMore(payload) || !lastSlug || lastSlug === startingAfter) {
      break;
    }
    startingAfter = lastSlug;
  }

  clients.sort((a, b) =>
    a.name.localeCompare(b.name, undefined, { sensitivity: "base" }),
  );
  return clients;
}

function normalizeStoredClient(value: unknown): IgnitionClientRecord | null {
  const fromApi = normalizeIgnitionClient(value);
  if (fromApi) {
    return fromApi;
  }

  const record = asRecord(value);
  if (!record || typeof record.id !== "string" || !record.id) {
    return null;
  }

  return {
    id: record.id,
    name: typeof record.name === "string" ? record.name : record.id,
    email: typeof record.email === "string" ? record.email : "",
    status: typeof record.status === "string" ? record.status : "",
    raw: asRecord(record.raw) ?? record,
  };
}

export async function readIgnitionClientsStore(): Promise<IgnitionClientsStore> {
  await ensureDataDir();
  try {
    const raw = await readFile(clientsPath(), "utf8");
    const parsed = JSON.parse(raw.replace(/^\uFEFF/, "")) as Partial<IgnitionClientsStore>;
    const clients = Array.isArray(parsed.clients)
      ? parsed.clients
          .map(normalizeStoredClient)
          .filter((item): item is IgnitionClientRecord => Boolean(item))
      : [];

    return {
      syncedAt: typeof parsed.syncedAt === "string" ? parsed.syncedAt : "",
      clients,
    };
  } catch (error) {
    if (isMissingFile(error)) {
      return { syncedAt: "", clients: [] };
    }
    throw error;
  }
}

export async function writeIgnitionClientsStore(
  store: IgnitionClientsStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(clientsPath(), store);
}

export async function syncIgnitionClients(): Promise<IgnitionClientsSyncResult> {
  const clients = await fetchAllIgnitionClients();
  const store: IgnitionClientsStore = {
    syncedAt: new Date().toISOString(),
    clients,
  };
  await writeIgnitionClientsStore(store);

  let newClientsFromAcceptedProposals: IgnitionAutoCreatedClientProfile[] = [];

  // Also sync accepted proposals and attach active PDF links to Client Profiles.
  let acceptedProposals: Awaited<
    ReturnType<typeof syncIgnitionProposals>
  >["proposals"] = [];
  try {
    const proposalsStore = await syncIgnitionProposals();
    acceptedProposals = proposalsStore.proposals;
    const matchClients = clients.map((client) => ({
      id: client.id,
      name: client.name,
      email: client.email,
    }));
    const profiles = await readClientProfiles();

    // Auto-add Client Info rows for accepted Ignition clients with no profile.
    const ensured = createProfilesForAcceptedIgnitionClients(
      profiles,
      matchClients,
      proposalsStore.proposals,
    );
    newClientsFromAcceptedProposals = ensured.created;

    const nextProfiles = applyActiveProposalsToProfiles(
      ensured.profiles,
      matchClients,
      proposalsStore.proposals,
    );
    await writeClientProfiles(nextProfiles);

    if (newClientsFromAcceptedProposals.length > 0) {
      await appendDashboardNotifications(
        newClientsFromAcceptedProposals.map((item) => ({
          kind: "ignition_new_client" as const,
          title: "New client added from Ignition",
          message: `${item.legalName} was added to Client Info after an accepted proposal${
            item.proposalReference ? ` (${item.proposalReference})` : ""
          }.`,
          clientName: item.legalName,
          clientProfileId: item.profileId,
        })),
      );

      for (const item of newClientsFromAcceptedProposals) {
        try {
          await appendHistory({
            employeeName: "Ignition sync",
            employeeEmail: "",
            project: item.legalName,
            entityType: "clientProfile",
            entityLabel: item.legalName,
            action: "created",
            summary: `Client profile auto-created from accepted Ignition proposal: ${item.legalName}`,
            details: [
              item.uniqueId ? `Unique ID: ${item.uniqueId}` : "",
              item.proposalReference
                ? `Proposal: ${item.proposalReference}`
                : "",
              item.ignitionClientSlug
                ? `Ignition client: ${item.ignitionClientSlug}`
                : "",
            ]
              .filter(Boolean)
              .join("\n"),
            comment: "Created automatically when an Ignition proposal was accepted.",
          });
        } catch (historyError) {
          console.error("Ignition new-client history failed:", historyError);
        }
      }
    }
  } catch (error) {
    // Client sync still succeeds if proposal sync fails; surface on next attempt.
    console.error("Ignition proposal sync failed:", error);
  }

  try {
    await syncIgnitionEndingSoon(
      acceptedProposals.length > 0 ? acceptedProposals : undefined,
    );
  } catch (error) {
    console.error("Ignition ending-soon sync failed:", error);
  }

  try {
    const clientNameBySlug: Record<string, string> = {};
    for (const client of clients) {
      clientNameBySlug[client.id] = client.name;
    }
    await syncIgnitionInvoices(clientNameBySlug);
  } catch (error) {
    console.error("Ignition invoice sync failed:", error);
  }

  return { ...store, newClientsFromAcceptedProposals };
}
