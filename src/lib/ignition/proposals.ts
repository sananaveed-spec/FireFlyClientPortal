import { copyFile, mkdir, readFile, unlink, writeFile } from "fs/promises";
import path from "path";
import {
  createEmptyClientProfileForm,
  generateUniqueId,
  normalizeClientProfile,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import { getDataDir } from "@/lib/data-store";
import { getIgnitionConfig } from "@/lib/ignition/config";
import { getValidAccessToken } from "@/lib/ignition/tokens";

export type IgnitionProposalRecord = {
  slug: string;
  clientSlug: string;
  clientName: string;
  name: string;
  referenceNumber: string;
  state: string;
  updatedAt: string;
  createdAt: string;
  /** Contract start date from Ignition when set (YYYY-MM-DD or ISO). */
  startDate: string;
  /** Minimum contract length in months, when provided. */
  minimumContractLength: number | null;
  /** Estimated end date (YYYY-MM-DD) from start + minimum term. */
  endsOn: string;
  viewUrl: string;
  pdfUrl: string;
};

export type IgnitionProposalsStore = {
  syncedAt: string;
  /** Month window used when this ending-soon list was built (if set). */
  windowMonths?: number;
  proposals: IgnitionProposalRecord[];
};

const IGNITION_APP_BASE = "https://go.ignitionapp.com";
/** Extra window beyond Ignition’s ending_soon flag (months from today). */
export const IGNITION_ENDING_SOON_MONTHS = 3;

function proposalsPath() {
  return path.join(getDataDir(), "ignition-proposals.json");
}

function endingSoonPath() {
  return path.join(getDataDir(), "ignition-ending-soon.json");
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

export function ignitionProposalViewUrl(
  clientSlug: string,
  proposalSlug: string,
): string {
  return `${IGNITION_APP_BASE}/clients/${encodeURIComponent(clientSlug)}/proposals/${encodeURIComponent(proposalSlug)}`;
}

export function ignitionProposalPdfUrl(
  clientSlug: string,
  proposalSlug: string,
): string {
  return `${ignitionProposalViewUrl(clientSlug, proposalSlug)}/pdf`;
}

export function normalizeMatchName(value: string): string {
  return value
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\b(llc|inc|ltd|corp|co|plc|pllc|pc)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/** Add whole months to a date; returns YYYY-MM-DD. */
export function addMonthsToDate(isoOrYmd: string, months: number): string | null {
  const raw = isoOrYmd.trim();
  if (!raw) {
    return null;
  }
  const date = new Date(raw.includes("T") ? raw : `${raw}T00:00:00.000Z`);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  date.setUTCMonth(date.getUTCMonth() + months);
  return date.toISOString().slice(0, 10);
}

function isYmdOrIsoDate(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}/.test(value.trim());
}

export function computeProposalEndDate(input: {
  startDate?: string;
  minimumContractLength?: number | null;
  updatedAt?: string;
  createdAt?: string;
}): string {
  const months = input.minimumContractLength;
  if (!months || months <= 0) {
    return "";
  }
  const candidates = [
    String(input.startDate ?? "").trim(),
    String(input.updatedAt ?? "").trim(),
    String(input.createdAt ?? "").trim(),
  ];
  const start = candidates.find((value) => isYmdOrIsoDate(value)) || "";
  return start ? addMonthsToDate(start, months) || "" : "";
}

export function isEndingWithinMonths(
  endsOn: string,
  months = IGNITION_ENDING_SOON_MONTHS,
  now: Date = new Date(),
): boolean {
  if (!/^\d{4}-\d{2}-\d{2}/.test(endsOn.trim())) {
    return false;
  }
  const end = Date.parse(endsOn.trim().slice(0, 10) + "T00:00:00.000Z");
  if (!Number.isFinite(end)) {
    return false;
  }
  const startOfToday = Date.UTC(
    now.getUTCFullYear(),
    now.getUTCMonth(),
    now.getUTCDate(),
  );
  const horizon = new Date(startOfToday);
  horizon.setUTCMonth(horizon.getUTCMonth() + months);
  return end >= startOfToday && end <= horizon.getTime();
}

export function normalizeIgnitionProposal(
  value: unknown,
): IgnitionProposalRecord | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const slug = pickString(record, ["slug", "id"]);
  const clientSlug = pickString(record, [
    "client_slug",
    "clientSlug",
    "client_id",
    "clientId",
  ]);
  if (!slug || !clientSlug) {
    return null;
  }
  const state = pickString(record, ["state", "status"]).toLowerCase();
  const updatedAt = pickString(record, ["updated_at", "updatedAt"]);
  const createdAt = pickString(record, ["created_at", "createdAt"]);
  // Ignition `start_on` is an enum ("date" | "acceptance"), not a calendar date.
  const startDateRaw = pickString(record, ["start_date", "startDate"]);
  const startDate = isYmdOrIsoDate(startDateRaw) ? startDateRaw : "";
  const minRaw = record.minimum_contract_length ?? record.minimumContractLength;
  const minParsed =
    typeof minRaw === "number"
      ? minRaw
      : Number.parseInt(String(minRaw ?? ""), 10);
  const minimumContractLength =
    Number.isFinite(minParsed) && minParsed > 0 ? minParsed : null;
  const endsOn =
    pickString(record, ["endsOn", "ends_on", "end_date", "endDate"]) ||
    computeProposalEndDate({
      startDate,
      minimumContractLength,
      updatedAt,
      createdAt,
    });
  return {
    slug,
    clientSlug,
    clientName: pickString(record, ["client_name", "clientName", "name"]),
    name: pickString(record, ["name", "title"]) || "Proposal",
    referenceNumber: pickString(record, [
      "reference_number",
      "referenceNumber",
      "reference",
    ]),
    state,
    updatedAt,
    createdAt,
    startDate,
    minimumContractLength,
    endsOn,
    viewUrl: ignitionProposalViewUrl(clientSlug, slug),
    pdfUrl: ignitionProposalPdfUrl(clientSlug, slug),
  };
}

function extractRows(payload: unknown): unknown[] {
  if (Array.isArray(payload)) {
    return payload;
  }
  const record = asRecord(payload);
  if (!record) {
    return [];
  }
  for (const key of ["data", "proposals", "items", "results", "records"]) {
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

async function fetchProposalsPage(
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
    const message =
      (record &&
        pickString(record, ["error_description", "message", "error"])) ||
      `Ignition proposals request failed (${response.status}).`;
    throw new Error(message);
  }
  return payload;
}

async function fetchIgnitionProposalsList(options: {
  state?: string;
  endingSoon?: boolean;
}): Promise<IgnitionProposalRecord[]> {
  const { accessToken } = await getValidAccessToken();
  const { apiBase } = getIgnitionConfig();

  const proposals: IgnitionProposalRecord[] = [];
  const seen = new Set<string>();
  let startingAfter = "";
  let pages = 0;

  while (pages < 50) {
    pages += 1;
    const url = new URL(`${apiBase}/proposals`);
    url.searchParams.set("limit", "100");
    if (options.state) {
      url.searchParams.set("state", options.state);
    }
    if (options.endingSoon) {
      url.searchParams.set("ending_soon", "true");
    }
    if (startingAfter) {
      url.searchParams.set("starting_after", startingAfter);
    }

    const payload = await fetchProposalsPage(url.toString(), accessToken);
    const rows = extractRows(payload);
    if (rows.length === 0) {
      break;
    }

    let lastSlug = "";
    for (const row of rows) {
      const proposal = normalizeIgnitionProposal(row);
      if (!proposal || seen.has(proposal.slug)) {
        continue;
      }
      seen.add(proposal.slug);
      proposals.push(proposal);
      lastSlug = proposal.slug;
    }

    if (!listHasMore(payload) || !lastSlug || lastSlug === startingAfter) {
      break;
    }
    startingAfter = lastSlug;
  }

  return proposals;
}

/** Accepted proposals are the active engagement PDFs in Ignition. */
export async function fetchAcceptedIgnitionProposals(): Promise<
  IgnitionProposalRecord[]
> {
  return fetchIgnitionProposalsList({ state: "accepted" });
}

/**
 * Ignition `ending_soon` proposals, plus accepted proposals ending within
 * {@link IGNITION_ENDING_SOON_MONTHS} months.
 */
export async function fetchEndingSoonIgnitionProposals(): Promise<
  IgnitionProposalRecord[]
> {
  const accepted = await fetchAcceptedIgnitionProposals();
  let apiEndingSoon: IgnitionProposalRecord[] = [];
  try {
    apiEndingSoon = await fetchIgnitionProposalsList({ endingSoon: true });
  } catch {
    // Accepted-list filter still applies if the ending_soon API call fails.
  }
  return filterEndingSoonProposals(accepted, apiEndingSoon);
}

function normalizeStoredProposal(value: unknown): IgnitionProposalRecord | null {
  const fromApi = normalizeIgnitionProposal(value);
  if (fromApi) {
    return fromApi;
  }
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const slug = pickString(record, ["slug"]);
  const clientSlug = pickString(record, ["clientSlug", "client_slug"]);
  if (!slug || !clientSlug) {
    return null;
  }
  const updatedAt = pickString(record, ["updatedAt", "updated_at"]);
  const createdAt = pickString(record, ["createdAt", "created_at"]);
  const startDateRaw = pickString(record, ["startDate", "start_date"]);
  const startDate = isYmdOrIsoDate(startDateRaw) ? startDateRaw : "";
  const minRaw = record.minimumContractLength ?? record.minimum_contract_length;
  const minParsed =
    typeof minRaw === "number"
      ? minRaw
      : Number.parseInt(String(minRaw ?? ""), 10);
  const minimumContractLength =
    Number.isFinite(minParsed) && minParsed > 0 ? minParsed : null;
  const endsOn =
    pickString(record, ["endsOn", "ends_on", "end_date", "endDate"]) ||
    computeProposalEndDate({
      startDate,
      minimumContractLength,
      updatedAt,
      createdAt,
    });
  return {
    slug,
    clientSlug,
    clientName: pickString(record, ["clientName", "client_name"]),
    name: pickString(record, ["name"]) || "Proposal",
    referenceNumber: pickString(record, [
      "referenceNumber",
      "reference_number",
    ]),
    state: pickString(record, ["state"]),
    updatedAt,
    createdAt,
    startDate,
    minimumContractLength,
    endsOn,
    viewUrl:
      pickString(record, ["viewUrl"]) ||
      ignitionProposalViewUrl(clientSlug, slug),
    pdfUrl:
      pickString(record, ["pdfUrl"]) ||
      ignitionProposalPdfUrl(clientSlug, slug),
  };
}

async function readProposalsStoreFile(
  filePath: string,
): Promise<IgnitionProposalsStore> {
  await ensureDataDir();
  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = JSON.parse(
      raw.replace(/^\uFEFF/, ""),
    ) as Partial<IgnitionProposalsStore>;
    const proposals = Array.isArray(parsed.proposals)
      ? parsed.proposals
          .map(normalizeStoredProposal)
          .filter((item): item is IgnitionProposalRecord => Boolean(item))
      : [];
    return {
      syncedAt: typeof parsed.syncedAt === "string" ? parsed.syncedAt : "",
      windowMonths:
        typeof parsed.windowMonths === "number" &&
        Number.isFinite(parsed.windowMonths)
          ? parsed.windowMonths
          : undefined,
      proposals,
    };
  } catch (error) {
    if (isMissingFile(error)) {
      return { syncedAt: "", proposals: [] };
    }
    throw error;
  }
}

export async function readIgnitionProposalsStore(): Promise<IgnitionProposalsStore> {
  return readProposalsStoreFile(proposalsPath());
}

export async function writeIgnitionProposalsStore(
  store: IgnitionProposalsStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(proposalsPath(), store);
}

export async function syncIgnitionProposals(): Promise<IgnitionProposalsStore> {
  const proposals = await fetchAcceptedIgnitionProposals();
  const store: IgnitionProposalsStore = {
    syncedAt: new Date().toISOString(),
    proposals,
  };
  await writeIgnitionProposalsStore(store);
  return store;
}

export async function readIgnitionEndingSoonStore(): Promise<IgnitionProposalsStore> {
  return readProposalsStoreFile(endingSoonPath());
}

export async function writeIgnitionEndingSoonStore(
  store: IgnitionProposalsStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(endingSoonPath(), store);
}

/**
 * Union of accepted + Ignition ending_soon candidates, kept only when the
 * computed contract end date falls in the next {@link IGNITION_ENDING_SOON_MONTHS}.
 * Rows with no computable end date are omitted (month window would be meaningless).
 */
export function filterEndingSoonProposals(
  accepted: IgnitionProposalRecord[],
  apiEndingSoon: IgnitionProposalRecord[] = [],
  months: number = IGNITION_ENDING_SOON_MONTHS,
): IgnitionProposalRecord[] {
  const bySlug = new Map<string, IgnitionProposalRecord>();

  for (const proposal of [...apiEndingSoon, ...accepted]) {
    if (!proposal.slug || bySlug.has(proposal.slug)) {
      continue;
    }
    if (
      proposal.endsOn &&
      isEndingWithinMonths(proposal.endsOn, months)
    ) {
      bySlug.set(proposal.slug, proposal);
    }
  }

  const filtered = Array.from(bySlug.values());
  filtered.sort((a, b) => {
    const aEnd = a.endsOn || "9999-12-31";
    const bEnd = b.endsOn || "9999-12-31";
    if (aEnd !== bEnd) {
      return aEnd.localeCompare(bEnd);
    }
    return a.clientName.localeCompare(b.clientName);
  });
  return filtered;
}

export async function syncIgnitionEndingSoon(
  acceptedProposals?: IgnitionProposalRecord[],
): Promise<IgnitionProposalsStore> {
  let proposals: IgnitionProposalRecord[];
  if (acceptedProposals) {
    let apiEndingSoon: IgnitionProposalRecord[] = [];
    try {
      apiEndingSoon = await fetchIgnitionProposalsList({ endingSoon: true });
    } catch {
      // Keep accepted-list window if ending_soon API fails.
    }
    proposals = filterEndingSoonProposals(acceptedProposals, apiEndingSoon);
  } else {
    proposals = await fetchEndingSoonIgnitionProposals();
  }
  const store: IgnitionProposalsStore = {
    syncedAt: new Date().toISOString(),
    windowMonths: IGNITION_ENDING_SOON_MONTHS,
    proposals,
  };
  await writeIgnitionEndingSoonStore(store);
  return store;
}

function proposalRecency(proposal: IgnitionProposalRecord): number {
  const stamp = Date.parse(proposal.updatedAt || proposal.createdAt);
  return Number.isFinite(stamp) ? stamp : 0;
}

/** One active (accepted) proposal per Ignition client — newest wins. */
export function pickActiveProposalByClient(
  proposals: IgnitionProposalRecord[],
): Map<string, IgnitionProposalRecord> {
  const byClient = new Map<string, IgnitionProposalRecord>();
  for (const proposal of proposals) {
    if (proposal.state && proposal.state !== "accepted") {
      continue;
    }
    const current = byClient.get(proposal.clientSlug);
    if (!current || proposalRecency(proposal) >= proposalRecency(current)) {
      byClient.set(proposal.clientSlug, proposal);
    }
  }
  return byClient;
}

export type IgnitionClientMatchSource = {
  id: string;
  name: string;
  email: string;
};

function profileEmails(profile: ClientProfileEntry): string[] {
  return [
    profile.ownerEmail,
    profile.pocEmail,
    profile.documentRequestEmail,
    profile.financialsEmail,
  ]
    .map((value) => value.trim().toLowerCase())
    .filter(Boolean);
}

function profileNames(profile: ClientProfileEntry): string[] {
  return [profile.legalName, profile.operatingName]
    .map((value) => normalizeMatchName(value))
    .filter(Boolean);
}

export function matchIgnitionClientToProfile(
  profile: ClientProfileEntry,
  clients: IgnitionClientMatchSource[],
): IgnitionClientMatchSource | null {
  if (profile.ignitionClientSlug.trim()) {
    const bySlug = clients.find(
      (client) => client.id === profile.ignitionClientSlug.trim(),
    );
    if (bySlug) {
      return bySlug;
    }
  }

  const names = new Set(profileNames(profile));
  for (const client of clients) {
    const clientName = normalizeMatchName(client.name);
    if (clientName && names.has(clientName)) {
      return client;
    }
  }

  const emails = new Set(profileEmails(profile));
  if (emails.size === 0) {
    return null;
  }
  for (const client of clients) {
    const email = client.email.trim().toLowerCase();
    if (email && emails.has(email)) {
      return client;
    }
  }
  return null;
}

export function applyActiveProposalToProfile(
  profile: ClientProfileEntry,
  clients: IgnitionClientMatchSource[],
  activeByClient: Map<string, IgnitionProposalRecord>,
): ClientProfileEntry {
  const matched = matchIgnitionClientToProfile(profile, clients);
  if (!matched) {
    return {
      ...profile,
      ignitionClientSlug: "",
      activeProposalSlug: "",
      activeProposalName: "",
      activeProposalReference: "",
      activeProposalViewUrl: "",
      activeProposalPdfUrl: "",
    };
  }

  const proposal = activeByClient.get(matched.id);
  if (!proposal) {
    return {
      ...profile,
      ignitionClientSlug: matched.id,
      activeProposalSlug: "",
      activeProposalName: "",
      activeProposalReference: "",
      activeProposalViewUrl: "",
      activeProposalPdfUrl: "",
    };
  }

  return {
    ...profile,
    ignitionClientSlug: matched.id,
    activeProposalSlug: proposal.slug,
    activeProposalName: proposal.name,
    activeProposalReference: proposal.referenceNumber || proposal.name,
    activeProposalViewUrl: proposal.viewUrl,
    activeProposalPdfUrl: proposal.pdfUrl,
  };
}

export function applyActiveProposalsToProfiles(
  profiles: ClientProfileEntry[],
  clients: IgnitionClientMatchSource[],
  proposals: IgnitionProposalRecord[],
): ClientProfileEntry[] {
  const activeByClient = pickActiveProposalByClient(proposals);
  return profiles.map((profile) =>
    applyActiveProposalToProfile(profile, clients, activeByClient),
  );
}

export type IgnitionAutoCreatedClientProfile = {
  profileId: string;
  uniqueId: string;
  legalName: string;
  ignitionClientSlug: string;
  proposalSlug: string;
  proposalReference: string;
};

function uniqueLegalName(
  preferred: string,
  existingLegalNames: Set<string>,
  ignitionClientId: string,
): string {
  const base = preferred.trim() || ignitionClientId;
  if (!existingLegalNames.has(base.toLowerCase())) {
    return base;
  }
  const suffix = ignitionClientId.replace(/^cli_/i, "").slice(-6) || "new";
  let candidate = `${base} (${suffix})`;
  let n = 2;
  while (existingLegalNames.has(candidate.toLowerCase())) {
    candidate = `${base} (${suffix}-${n})`;
    n += 1;
  }
  return candidate;
}

/**
 * For Ignition clients with an accepted proposal and no matching Client Info
 * profile, create a new Client Profile and attach the active proposal.
 */
export function createProfilesForAcceptedIgnitionClients(
  profiles: ClientProfileEntry[],
  clients: IgnitionClientMatchSource[],
  proposals: IgnitionProposalRecord[],
): {
  profiles: ClientProfileEntry[];
  created: IgnitionAutoCreatedClientProfile[];
} {
  const activeByClient = pickActiveProposalByClient(proposals);
  if (activeByClient.size === 0) {
    return { profiles, created: [] };
  }

  const next = [...profiles];
  const created: IgnitionAutoCreatedClientProfile[] = [];
  const existingLegalNames = new Set(
    next
      .map((profile) => profile.legalName.trim().toLowerCase())
      .filter(Boolean),
  );
  const existingUniqueIds = next.map((profile) => profile.uniqueId);

  for (const client of clients) {
    if (!activeByClient.has(client.id)) {
      continue;
    }
    const alreadyLinked = next.some(
      (profile) =>
        matchIgnitionClientToProfile(profile, [client])?.id === client.id,
    );
    if (alreadyLinked) {
      continue;
    }

    const proposal = activeByClient.get(client.id)!;
    const legalName = uniqueLegalName(
      client.name || proposal.clientName,
      existingLegalNames,
      client.id,
    );
    existingLegalNames.add(legalName.toLowerCase());

    const form = createEmptyClientProfileForm(existingUniqueIds, "US");
    const uniqueId = form.uniqueId || generateUniqueId(existingUniqueIds, "US");
    existingUniqueIds.push(uniqueId);

    const email = client.email.trim();
    const profileId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const profile = normalizeClientProfile({
      id: profileId,
      ...form,
      uniqueId,
      legalName,
      operatingName: client.name.trim() || legalName,
      ownerEmail: email,
      pocEmail: email,
      ignitionClientSlug: client.id,
      activeProposalSlug: proposal.slug,
      activeProposalName: proposal.name,
      activeProposalReference: proposal.referenceNumber || proposal.name,
      activeProposalViewUrl: proposal.viewUrl,
      activeProposalPdfUrl: proposal.pdfUrl,
      createdAt: new Date().toISOString(),
    });

    next.unshift(profile);
    created.push({
      profileId: profile.id,
      uniqueId: profile.uniqueId,
      legalName: profile.legalName,
      ignitionClientSlug: client.id,
      proposalSlug: proposal.slug,
      proposalReference:
        proposal.referenceNumber || proposal.name || proposal.slug,
    });
  }

  return { profiles: next, created };
}

export type IgnitionProposalServiceView = {
  name: string;
  descriptionHtml: string;
  amountLabel: string;
  billingMode: string;
};

export type IgnitionProposalView = {
  slug: string;
  name: string;
  clientSlug: string;
  clientName: string;
  referenceNumber: string;
  state: string;
  personalisedMessageHtml: string;
  nextStepsMessageHtml: string;
  contractTerm: string;
  minimumContractLength: number | null;
  services: IgnitionProposalServiceView[];
  /** Ignition deep links (optional; PDF bytes are not available via Core API). */
  ignitionViewUrl: string;
  ignitionPdfUrl: string;
};

function sanitizeIgnitionHtml(html: string): string {
  return html
    .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, "")
    .replace(/\son\w+\s*=\s*(['"]).*?\1/gi, "")
    .replace(/\son\w+\s*=\s*[^\s>]+/gi, "");
}

function formatCents(amountCents: unknown, currency: string): string {
  const cents =
    typeof amountCents === "number" && Number.isFinite(amountCents)
      ? amountCents
      : Number(amountCents);
  if (!Number.isFinite(cents)) {
    return "";
  }
  const amount = cents / 100;
  try {
    return new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: currency || "USD",
    }).format(amount);
  } catch {
    return `${currency || "USD"} ${amount.toFixed(2)}`;
  }
}

function extractServices(
  document: Record<string, unknown>,
): IgnitionProposalServiceView[] {
  const services: IgnitionProposalServiceView[] = [];
  const options = Array.isArray(document.options) ? document.options : [];
  for (const option of options) {
    const optionRecord = asRecord(option);
    if (!optionRecord) {
      continue;
    }
    const projects = Array.isArray(optionRecord.projects)
      ? optionRecord.projects
      : [];
    for (const project of projects) {
      const projectRecord = asRecord(project);
      if (!projectRecord) {
        continue;
      }
      const groups = Array.isArray(projectRecord.service_groups)
        ? projectRecord.service_groups
        : [];
      for (const group of groups) {
        const groupRecord = asRecord(group);
        if (!groupRecord) {
          continue;
        }
        const billingMode = pickString(groupRecord, [
          "configured_as",
          "billing_mode",
        ]);
        const proposed = Array.isArray(groupRecord.proposed_services)
          ? groupRecord.proposed_services
          : [];
        for (const service of proposed) {
          const serviceRecord = asRecord(service);
          if (!serviceRecord) {
            continue;
          }
          const priceRule = asRecord(serviceRecord.price_rule);
          const amountLabel = priceRule
            ? formatCents(
                priceRule.amount_cents,
                pickString(priceRule, ["currency"]) || "USD",
              )
            : "";
          services.push({
            name:
              pickString(serviceRecord, ["name", "billing_name"]) || "Service",
            descriptionHtml: sanitizeIgnitionHtml(
              pickString(serviceRecord, ["description"]),
            ),
            amountLabel,
            billingMode,
          });
        }
      }
    }
  }
  return services;
}

export async function fetchIgnitionProposalView(
  proposalSlug: string,
): Promise<IgnitionProposalView> {
  const slug = proposalSlug.trim();
  if (!slug) {
    throw new Error("Proposal slug is required.");
  }

  const { accessToken } = await getValidAccessToken();
  const { apiBase } = getIgnitionConfig();
  const store = await readIgnitionProposalsStore();
  const cached = store.proposals.find((item) => item.slug === slug);

  const [metaRes, docRes] = await Promise.all([
    fetch(`${apiBase}/proposals/${encodeURIComponent(slug)}`, {
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
      cache: "no-store",
    }),
    fetch(`${apiBase}/proposals/${encodeURIComponent(slug)}/document`, {
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
      cache: "no-store",
    }),
  ]);

  const metaPayload = await metaRes.json().catch(() => ({}));
  const docPayload = await docRes.json().catch(() => ({}));
  if (!metaRes.ok) {
    const record = asRecord(metaPayload);
    throw new Error(
      (record && pickString(record, ["error_description", "message", "error"])) ||
        `Failed to load proposal (${metaRes.status}).`,
    );
  }
  if (!docRes.ok) {
    const record = asRecord(docPayload);
    throw new Error(
      (record && pickString(record, ["error_description", "message", "error"])) ||
        `Failed to load proposal document (${docRes.status}).`,
    );
  }

  const meta = asRecord(metaPayload) ?? {};
  const docRoot = asRecord(docPayload) ?? {};
  const document = asRecord(docRoot.document) ?? {};
  const clientSlug =
    pickString(meta, ["client_slug", "clientSlug"]) ||
    pickString(document, ["client_slug", "clientSlug"]) ||
    cached?.clientSlug ||
    "";
  const proposalName =
    pickString(document, ["name"]) ||
    pickString(meta, ["name"]) ||
    cached?.name ||
    "Proposal";

  const minLengthRaw = document.minimum_contract_length;
  const minimumContractLength =
    typeof minLengthRaw === "number" && Number.isFinite(minLengthRaw)
      ? minLengthRaw
      : null;

  return {
    slug,
    name: proposalName,
    clientSlug,
    clientName:
      pickString(meta, ["client_name", "clientName"]) ||
      cached?.clientName ||
      "",
    referenceNumber:
      pickString(meta, ["reference_number", "referenceNumber"]) ||
      cached?.referenceNumber ||
      "",
    state: pickString(meta, ["state", "status"]) || cached?.state || "",
    personalisedMessageHtml: sanitizeIgnitionHtml(
      pickString(document, ["personalised_message"]),
    ),
    nextStepsMessageHtml: sanitizeIgnitionHtml(
      pickString(document, ["next_steps_message"]),
    ),
    contractTerm: pickString(document, ["contract_term"]),
    minimumContractLength,
    services: extractServices(document),
    ignitionViewUrl:
      cached?.viewUrl ||
      (clientSlug ? ignitionProposalViewUrl(clientSlug, slug) : ""),
    ignitionPdfUrl:
      cached?.pdfUrl ||
      (clientSlug ? ignitionProposalPdfUrl(clientSlug, slug) : ""),
  };
}
