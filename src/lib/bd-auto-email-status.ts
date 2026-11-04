import { currentMonthKey } from "@/lib/client-request-lists";

/** Schedule ids that auto-send without supervisor click. */
export const BD_AUTO_SEND_SCHEDULE_IDS = ["bd1", "bd7", "bd8"] as const;

export type BdAutoSendScheduleId = (typeof BD_AUTO_SEND_SCHEDULE_IDS)[number];

export function isBdAutoSendScheduleId(
  value: string,
): value is BdAutoSendScheduleId {
  return (BD_AUTO_SEND_SCHEDULE_IDS as readonly string[]).includes(value);
}

export type BdAutoEmailStatus = {
  clientId: string;
  scheduleId: BdAutoSendScheduleId;
  monthKey: string;
  emailSent: boolean;
  /** ISO timestamp when send completed (or was marked). */
  lastSentAt: string;
  /** Supervisor mailbox used as From. */
  fromEmail: string;
  /** How it was recorded. */
  source: "auto-cron" | "manual-schedule" | "manual-mark";
  lastError: string;
};

export type BdAutoEmailStatusStore = {
  statuses: BdAutoEmailStatus[];
};

export function createEmptyBdAutoEmailStatusStore(): BdAutoEmailStatusStore {
  return { statuses: [] };
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

export function normalizeBdAutoEmailStatus(
  value: unknown,
): BdAutoEmailStatus | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const clientId = String(record.clientId ?? "").trim();
  const scheduleIdRaw = String(record.scheduleId ?? "").trim();
  const monthKey = String(record.monthKey ?? "").trim();
  if (!clientId || !monthKey || !isBdAutoSendScheduleId(scheduleIdRaw)) {
    return null;
  }
  const sourceRaw = String(record.source ?? "").trim();
  const source: BdAutoEmailStatus["source"] =
    sourceRaw === "manual-schedule" || sourceRaw === "manual-mark"
      ? sourceRaw
      : "auto-cron";
  return {
    clientId,
    scheduleId: scheduleIdRaw,
    monthKey,
    emailSent: Boolean(record.emailSent),
    lastSentAt:
      typeof record.lastSentAt === "string" && record.lastSentAt.trim()
        ? record.lastSentAt.trim()
        : typeof record.emailSentAt === "string" && record.emailSentAt.trim()
          ? record.emailSentAt.trim()
          : "",
    fromEmail: String(record.fromEmail ?? "").trim(),
    source,
    lastError: String(record.lastError ?? "").trim(),
  };
}

/** Keep current month (+ prior month briefly for late PST catch-up display). */
export function pruneBdAutoEmailStatus(
  store: BdAutoEmailStatusStore,
  monthKey: string = currentMonthKey(),
): BdAutoEmailStatusStore {
  const [yearRaw, monthRaw] = monthKey.split("-");
  const year = Number(yearRaw);
  const month = Number(monthRaw);
  let priorKey = "";
  if (Number.isFinite(year) && Number.isFinite(month) && month >= 1) {
    if (month === 1) {
      priorKey = `${year - 1}-12`;
    } else {
      priorKey = `${year}-${String(month - 1).padStart(2, "0")}`;
    }
  }
  return {
    statuses: store.statuses.filter(
      (item) => item.monthKey === monthKey || item.monthKey === priorKey,
    ),
  };
}

export function normalizeBdAutoEmailStatusStore(
  value: unknown,
  monthKey: string = currentMonthKey(),
): BdAutoEmailStatusStore {
  const record = asRecord(value);
  const list = Array.isArray(record?.statuses)
    ? record.statuses
    : Array.isArray(value)
      ? value
      : [];
  const statuses = list
    .map(normalizeBdAutoEmailStatus)
    .filter((item): item is BdAutoEmailStatus => Boolean(item));
  return pruneBdAutoEmailStatus({ statuses }, monthKey);
}

export function getBdAutoEmailStatus(
  store: BdAutoEmailStatusStore,
  clientId: string,
  scheduleId: string,
  monthKey: string = currentMonthKey(),
): BdAutoEmailStatus | null {
  if (!isBdAutoSendScheduleId(scheduleId)) {
    return null;
  }
  return (
    store.statuses.find(
      (item) =>
        item.clientId === clientId &&
        item.scheduleId === scheduleId &&
        item.monthKey === monthKey,
    ) ?? null
  );
}

export function markBdAutoEmailSent(
  store: BdAutoEmailStatusStore,
  input: {
    clientId: string;
    scheduleId: BdAutoSendScheduleId;
    monthKey?: string;
    lastSentAt?: string;
    fromEmail?: string;
    source?: BdAutoEmailStatus["source"];
  },
): BdAutoEmailStatusStore {
  const monthKey = input.monthKey?.trim() || currentMonthKey();
  const lastSentAt = input.lastSentAt?.trim() || new Date().toISOString();
  const next: BdAutoEmailStatus = {
    clientId: input.clientId.trim(),
    scheduleId: input.scheduleId,
    monthKey,
    emailSent: true,
    lastSentAt,
    fromEmail: input.fromEmail?.trim() || "",
    source: input.source ?? "auto-cron",
    lastError: "",
  };
  const statuses = store.statuses.filter(
    (item) =>
      !(
        item.clientId === next.clientId &&
        item.scheduleId === next.scheduleId &&
        item.monthKey === next.monthKey
      ),
  );
  statuses.push(next);
  return pruneBdAutoEmailStatus({ statuses }, monthKey);
}

export function markBdAutoEmailError(
  store: BdAutoEmailStatusStore,
  input: {
    clientId: string;
    scheduleId: BdAutoSendScheduleId;
    monthKey?: string;
    lastError: string;
    fromEmail?: string;
  },
): BdAutoEmailStatusStore {
  const monthKey = input.monthKey?.trim() || currentMonthKey();
  const existing = getBdAutoEmailStatus(
    store,
    input.clientId,
    input.scheduleId,
    monthKey,
  );
  const next: BdAutoEmailStatus = {
    clientId: input.clientId.trim(),
    scheduleId: input.scheduleId,
    monthKey,
    emailSent: existing?.emailSent ?? false,
    lastSentAt: existing?.lastSentAt ?? "",
    fromEmail: input.fromEmail?.trim() || existing?.fromEmail || "",
    source: existing?.source ?? "auto-cron",
    lastError: input.lastError.trim(),
  };
  const statuses = store.statuses.filter(
    (item) =>
      !(
        item.clientId === next.clientId &&
        item.scheduleId === next.scheduleId &&
        item.monthKey === next.monthKey
      ),
  );
  statuses.push(next);
  return pruneBdAutoEmailStatus({ statuses }, monthKey);
}
