import { copyFile, mkdir, readFile, unlink, writeFile } from "fs/promises";
import path from "path";
import { getDataDir } from "@/lib/data-store";
import { getIgnitionConfig } from "@/lib/ignition/config";
import { getValidAccessToken } from "@/lib/ignition/tokens";

export type IgnitionInvoiceRecord = {
  slug: string;
  number: string;
  clientSlug: string;
  clientName: string;
  state: string;
  paymentState: string;
  amountCents: number;
  taxAmountCents: number;
  currency: string;
  /** Invoice issue date (YYYY-MM-DD). */
  date: string;
  /** Due date (YYYY-MM-DD). */
  dueOn: string;
  /** Paid date when known (YYYY-MM-DD or ISO). */
  paidOn: string;
  createdAt: string;
  updatedAt: string;
  raw: Record<string, unknown>;
};

export type IgnitionInvoicesStore = {
  syncedAt: string;
  invoices: IgnitionInvoiceRecord[];
};

/** One row per client for the Invoices tab. */
export type IgnitionClientInvoiceSummary = {
  clientSlug: string;
  clientName: string;
  lastPaidOn: string;
  lastPaidAmountCents: number | null;
  lastPaidCurrency: string;
  lastPaidInvoiceNumber: string;
  /** Oldest unpaid invoice due before today (overdue outstanding). */
  previousDuesOn: string;
  /** Sum of all overdue unpaid invoice amounts. */
  previousDuesAmountCents: number | null;
  previousDuesCurrency: string;
  previousDuesInvoiceNumber: string;
  previousDuesCount: number;
  /** Soonest unpaid invoice due today or later. */
  nextUpcomingOn: string;
  nextUpcomingAmountCents: number | null;
  nextUpcomingCurrency: string;
  nextUpcomingInvoiceNumber: string;
};

function invoicesPath() {
  return path.join(getDataDir(), "ignition-invoices.json");
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

function pickCents(record: Record<string, unknown>, keys: string[]): number {
  for (const key of keys) {
    const value = record[key];
    if (typeof value === "number" && Number.isFinite(value)) {
      return Math.round(value);
    }
    if (typeof value === "string" && value.trim()) {
      const parsed = Number(value);
      if (Number.isFinite(parsed)) {
        return Math.round(parsed);
      }
    }
  }
  return 0;
}

function dateKey(value: string): string {
  const raw = value.trim();
  if (!raw) {
    return "";
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) {
    return raw.slice(0, 10);
  }
  const parsed = Date.parse(raw);
  if (!Number.isFinite(parsed)) {
    return "";
  }
  return new Date(parsed).toISOString().slice(0, 10);
}

function isPaidState(paymentState: string): boolean {
  const state = paymentState.toLowerCase();
  return state === "paid" || state === "paid_in_full" || state === "complete";
}

function isOpenForNextDue(invoice: IgnitionInvoiceRecord): boolean {
  const state = invoice.state.toLowerCase();
  if (
    state === "voided" ||
    state === "void" ||
    state === "draft" ||
    state === "archived"
  ) {
    return false;
  }
  const payment = invoice.paymentState.toLowerCase();
  if (isPaidState(payment)) {
    return false;
  }
  // unpaid / partially_paid / overdue-style states stay in next-due pool
  return Boolean(invoice.dueOn);
}

export function formatIgnitionInvoiceAmount(
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

export function normalizeIgnitionInvoice(
  value: unknown,
  clientNameBySlug: Record<string, string> = {},
): IgnitionInvoiceRecord | null {
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

  const nestedClient = asRecord(record.client);
  const clientName =
    pickString(record, ["client_name", "clientName"]) ||
    (nestedClient
      ? pickString(nestedClient, ["name", "client_name", "clientName"])
      : "") ||
    clientNameBySlug[clientSlug] ||
    "";

  const paymentState = pickString(record, [
    "payment_state",
    "paymentState",
    "payment_status",
    "paymentStatus",
  ]).toLowerCase();

  const paidOn =
    pickString(record, [
      "paid_on",
      "paidOn",
      "payment_date",
      "paymentDate",
      "paid_at",
      "paidAt",
      "collected_on",
      "collectedOn",
    ]) ||
    (isPaidState(paymentState)
      ? pickString(record, ["updated_at", "updatedAt", "date"])
      : "");

  return {
    slug,
    number: pickString(record, [
      "number",
      "reference_number",
      "referenceNumber",
      "reference",
    ]),
    clientSlug,
    clientName,
    state: pickString(record, ["state", "status"]).toLowerCase(),
    paymentState,
    amountCents: pickCents(record, [
      "amount_cents",
      "amountCents",
      "total_cents",
      "totalCents",
    ]),
    taxAmountCents: pickCents(record, [
      "tax_amount_cents",
      "taxAmountCents",
    ]),
    currency: pickString(record, ["currency"]) || "USD",
    date: dateKey(pickString(record, ["date", "issued_on", "issuedOn"])),
    dueOn: dateKey(pickString(record, ["due_on", "dueOn", "due_date", "dueDate"])),
    paidOn: dateKey(paidOn),
    createdAt: pickString(record, ["created_at", "createdAt"]),
    updatedAt: pickString(record, ["updated_at", "updatedAt"]),
    raw: record,
  };
}

function extractInvoiceRows(payload: unknown): unknown[] {
  if (Array.isArray(payload)) {
    return payload;
  }
  const record = asRecord(payload);
  if (!record) {
    return [];
  }
  for (const key of ["data", "invoices", "items", "results", "records"]) {
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

async function fetchInvoicesPage(
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
    const message =
      (errorObj && pickString(errorObj, ["message", "code"])) ||
      (record &&
        pickString(record, ["error_description", "message", "error"])) ||
      `Ignition invoices request failed (${response.status}).`;
    throw new Error(message);
  }
  return payload;
}

export async function fetchAllIgnitionInvoices(
  clientNameBySlug: Record<string, string> = {},
): Promise<IgnitionInvoiceRecord[]> {
  const { accessToken } = await getValidAccessToken();
  const { apiBase } = getIgnitionConfig();

  const invoices: IgnitionInvoiceRecord[] = [];
  const seen = new Set<string>();
  let startingAfter = "";
  let pages = 0;

  while (pages < 50) {
    pages += 1;
    const url = new URL(`${apiBase}/invoices`);
    url.searchParams.set("limit", "100");
    if (startingAfter) {
      url.searchParams.set("starting_after", startingAfter);
    }

    const payload = await fetchInvoicesPage(url.toString(), accessToken);
    const rows = extractInvoiceRows(payload);
    if (rows.length === 0) {
      break;
    }

    let lastSlug = "";
    for (const row of rows) {
      const invoice = normalizeIgnitionInvoice(row, clientNameBySlug);
      if (!invoice || seen.has(invoice.slug)) {
        continue;
      }
      seen.add(invoice.slug);
      invoices.push(invoice);
      lastSlug = invoice.slug;
    }

    if (!listHasMore(payload) || !lastSlug || lastSlug === startingAfter) {
      break;
    }
    startingAfter = lastSlug;
  }

  invoices.sort((a, b) => {
    const byDue = (b.dueOn || b.date).localeCompare(a.dueOn || a.date);
    if (byDue !== 0) {
      return byDue;
    }
    return a.number.localeCompare(b.number, undefined, {
      sensitivity: "base",
    });
  });

  return invoices;
}

function normalizeStoredInvoice(
  value: unknown,
): IgnitionInvoiceRecord | null {
  const fromApi = normalizeIgnitionInvoice(value);
  if (fromApi && !asRecord(value)?.raw) {
    return fromApi;
  }

  const record = asRecord(value);
  if (!record) {
    return null;
  }
  if (typeof record.slug === "string" && record.slug && record.raw) {
    const nested = normalizeIgnitionInvoice(record.raw, {
      [String(record.clientSlug ?? "")]: String(record.clientName ?? ""),
    });
    if (nested) {
      return {
        ...nested,
        clientName:
          nested.clientName ||
          (typeof record.clientName === "string" ? record.clientName : ""),
      };
    }
  }
  return normalizeIgnitionInvoice(record);
}

export async function readIgnitionInvoicesStore(): Promise<IgnitionInvoicesStore> {
  await ensureDataDir();
  try {
    const raw = await readFile(invoicesPath(), "utf8");
    const parsed = JSON.parse(raw.replace(/^\uFEFF/, "")) as Partial<IgnitionInvoicesStore>;
    const invoices = Array.isArray(parsed.invoices)
      ? parsed.invoices
          .map(normalizeStoredInvoice)
          .filter((item): item is IgnitionInvoiceRecord => Boolean(item))
      : [];
    return {
      syncedAt: typeof parsed.syncedAt === "string" ? parsed.syncedAt : "",
      invoices,
    };
  } catch (error) {
    if (isMissingFile(error)) {
      return { syncedAt: "", invoices: [] };
    }
    throw error;
  }
}

export async function writeIgnitionInvoicesStore(
  store: IgnitionInvoicesStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(invoicesPath(), store);
}

export async function syncIgnitionInvoices(
  clientNameBySlug: Record<string, string> = {},
): Promise<IgnitionInvoicesStore> {
  const invoices = await fetchAllIgnitionInvoices(clientNameBySlug);
  const store: IgnitionInvoicesStore = {
    syncedAt: new Date().toISOString(),
    invoices,
  };
  await writeIgnitionInvoicesStore(store);
  return store;
}

function todayYmd(now: Date = new Date()): string {
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function pickSoonestOpen(
  invoices: IgnitionInvoiceRecord[],
  predicate: (invoice: IgnitionInvoiceRecord) => boolean,
): IgnitionInvoiceRecord | null {
  let best: IgnitionInvoiceRecord | null = null;
  for (const invoice of invoices) {
    if (!isOpenForNextDue(invoice) || !predicate(invoice)) {
      continue;
    }
    if (!best) {
      best = invoice;
      continue;
    }
    if (invoice.dueOn < best.dueOn) {
      best = invoice;
    } else if (
      invoice.dueOn === best.dueOn &&
      invoice.number.localeCompare(best.number) < 0
    ) {
      best = invoice;
    }
  }
  return best;
}

/**
 * Per-client invoice summary from synced Ignition invoices.
 * - Last paid = most recent paid invoice
 * - Previous dues = unpaid invoices due before today (overdue); amount is the total
 * - Next upcoming bill = soonest unpaid invoice due today or later
 */
export function buildIgnitionClientInvoiceSummaries(
  invoices: IgnitionInvoiceRecord[],
  clientNameBySlug: Record<string, string> = {},
  now: Date = new Date(),
): IgnitionClientInvoiceSummary[] {
  const today = todayYmd(now);
  const byClient = new Map<string, IgnitionInvoiceRecord[]>();
  for (const invoice of invoices) {
    const slug = invoice.clientSlug.trim();
    if (!slug) {
      continue;
    }
    const list = byClient.get(slug) ?? [];
    list.push(invoice);
    byClient.set(slug, list);
  }

  const summaries: IgnitionClientInvoiceSummary[] = [];
  for (const [clientSlug, list] of byClient) {
    const clientName =
      clientNameBySlug[clientSlug] ||
      list.find((item) => item.clientName)?.clientName ||
      clientSlug;

    let lastPaid: IgnitionInvoiceRecord | null = null;
    for (const invoice of list) {
      if (!isPaidState(invoice.paymentState)) {
        continue;
      }
      const paidKey = invoice.paidOn || invoice.date || invoice.updatedAt;
      if (!paidKey) {
        continue;
      }
      if (!lastPaid) {
        lastPaid = invoice;
        continue;
      }
      const prevKey =
        lastPaid.paidOn || lastPaid.date || lastPaid.updatedAt;
      if (dateKey(paidKey) > dateKey(prevKey)) {
        lastPaid = invoice;
      }
    }

    const overdue = list.filter(
      (invoice) => isOpenForNextDue(invoice) && invoice.dueOn < today,
    );
    const oldestOverdue = pickSoonestOpen(list, (invoice) => invoice.dueOn < today);
    const nextUpcoming = pickSoonestOpen(
      list,
      (invoice) => invoice.dueOn >= today,
    );

    let previousDuesAmount: number | null = null;
    let previousDuesCurrency = "USD";
    if (overdue.length > 0) {
      previousDuesAmount = overdue.reduce(
        (sum, invoice) => sum + invoice.amountCents,
        0,
      );
      previousDuesCurrency = oldestOverdue?.currency || overdue[0].currency || "USD";
    }

    const previousDuesInvoiceNumber =
      overdue.length === 0
        ? ""
        : overdue.length === 1
          ? overdue[0].number
          : `${oldestOverdue?.number || overdue[0].number} +${overdue.length - 1}`;

    summaries.push({
      clientSlug,
      clientName,
      lastPaidOn: lastPaid
        ? lastPaid.paidOn || lastPaid.date || dateKey(lastPaid.updatedAt)
        : "",
      lastPaidAmountCents: lastPaid ? lastPaid.amountCents : null,
      lastPaidCurrency: lastPaid?.currency || "USD",
      lastPaidInvoiceNumber: lastPaid?.number || "",
      previousDuesOn: oldestOverdue?.dueOn || "",
      previousDuesAmountCents: previousDuesAmount,
      previousDuesCurrency,
      previousDuesInvoiceNumber,
      previousDuesCount: overdue.length,
      nextUpcomingOn: nextUpcoming?.dueOn || "",
      nextUpcomingAmountCents: nextUpcoming ? nextUpcoming.amountCents : null,
      nextUpcomingCurrency: nextUpcoming?.currency || "USD",
      nextUpcomingInvoiceNumber: nextUpcoming?.number || "",
    });
  }

  summaries.sort((a, b) =>
    a.clientName.localeCompare(b.clientName, undefined, {
      sensitivity: "base",
    }),
  );
  return summaries;
}
