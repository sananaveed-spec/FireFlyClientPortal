import { currentMonthKey } from "@/lib/client-request-lists";

/** Schedule ids that auto-send without supervisor click. */
export const BD_AUTO_SEND_SCHEDULE_IDS = ["bd1", "bd7", "bd8"] as const;

export type BdAutoSendScheduleId = (typeof BD_AUTO_SEND_SCHEDULE_IDS)[number];

export function isBdAutoSendScheduleId(
  value: string,
): value is BdAutoSendScheduleId {
  return (BD_AUTO_SEND_SCHEDULE_IDS as readonly string[]).includes(value);
}

export type BdAutoEmailAttempt = {
  /** ISO timestamp when this cron/check ran. */
  at: string;
  outcome: "sent" | "error";
  /** Human message (error text, or short success note). */
  message: string;
  fromEmail: string;
};

export type BdAutoEmailStatus = {
  clientId: string;
  scheduleId: BdAutoSendScheduleId;
  monthKey: string;
  emailSent: boolean;
  /** ISO timestamp when send completed (or was marked). */
  lastSentAt: string;
  /** Supervisor mailbox used as From. */
  fromEmail: string;
  /** How the successful send was recorded. */
  source: "auto-cron" | "manual-schedule" | "manual-mark";
  lastError: string;
  /** Cron/send tries after send time (newest last). Not used for "already sent" skips. */
  attempts: BdAutoEmailAttempt[];
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

function normalizeAttempt(value: unknown): BdAutoEmailAttempt | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const at = String(record.at ?? "").trim();
  const outcome = String(record.outcome ?? "").trim();
  if (!at || (outcome !== "sent" && outcome !== "error")) {
    return null;
  }
  return {
    at,
    outcome,
    message: String(record.message ?? "").trim(),
    fromEmail: String(record.fromEmail ?? "").trim(),
  };
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
  const lastSentAt =
    typeof record.lastSentAt === "string" && record.lastSentAt.trim()
      ? record.lastSentAt.trim()
      : typeof record.emailSentAt === "string" && record.emailSentAt.trim()
        ? record.emailSentAt.trim()
        : "";
  const lastError = String(record.lastError ?? "").trim();
  let attempts = Array.isArray(record.attempts)
    ? record.attempts
        .map(normalizeAttempt)
        .filter((item): item is BdAutoEmailAttempt => Boolean(item))
    : [];
  // Migrate legacy lastError / lastSentAt into attempts when empty.
  if (attempts.length === 0) {
    if (Boolean(record.emailSent) && lastSentAt) {
      attempts = [
        {
          at: lastSentAt,
          outcome: "sent",
          message: "Sent",
          fromEmail: String(record.fromEmail ?? "").trim(),
        },
      ];
    } else if (lastError) {
      attempts = [
        {
          at: lastSentAt || new Date().toISOString(),
          outcome: "error",
          message: lastError,
          fromEmail: String(record.fromEmail ?? "").trim(),
        },
      ];
    }
  }
  return {
    clientId,
    scheduleId: scheduleIdRaw,
    monthKey,
    emailSent: Boolean(record.emailSent),
    lastSentAt,
    fromEmail: String(record.fromEmail ?? "").trim(),
    source,
    lastError,
    attempts,
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

function upsertStatus(
  store: BdAutoEmailStatusStore,
  next: BdAutoEmailStatus,
  monthKey: string,
): BdAutoEmailStatusStore {
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

function appendAttempt(
  existing: BdAutoEmailStatus | null,
  attempt: BdAutoEmailAttempt,
): BdAutoEmailAttempt[] {
  const prior = existing?.attempts ?? [];
  return [...prior, attempt].slice(-20);
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
    attemptMessage?: string;
  },
): BdAutoEmailStatusStore {
  const monthKey = input.monthKey?.trim() || currentMonthKey();
  const lastSentAt = input.lastSentAt?.trim() || new Date().toISOString();
  const fromEmail = input.fromEmail?.trim() || "";
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
    emailSent: true,
    lastSentAt,
    fromEmail: fromEmail || existing?.fromEmail || "",
    source: input.source ?? "auto-cron",
    lastError: "",
    attempts: appendAttempt(existing, {
      at: lastSentAt,
      outcome: "sent",
      message: input.attemptMessage?.trim() || "Sent",
      fromEmail: fromEmail || existing?.fromEmail || "",
    }),
  };
  return upsertStatus(store, next, monthKey);
}

export function markBdAutoEmailError(
  store: BdAutoEmailStatusStore,
  input: {
    clientId: string;
    scheduleId: BdAutoSendScheduleId;
    monthKey?: string;
    lastError: string;
    fromEmail?: string;
    attemptedAt?: string;
  },
): BdAutoEmailStatusStore {
  const monthKey = input.monthKey?.trim() || currentMonthKey();
  const attemptedAt = input.attemptedAt?.trim() || new Date().toISOString();
  const existing = getBdAutoEmailStatus(
    store,
    input.clientId,
    input.scheduleId,
    monthKey,
  );
  const fromEmail = input.fromEmail?.trim() || existing?.fromEmail || "";
  const lastError = input.lastError.trim();
  const next: BdAutoEmailStatus = {
    clientId: input.clientId.trim(),
    scheduleId: input.scheduleId,
    monthKey,
    emailSent: existing?.emailSent ?? false,
    lastSentAt: existing?.lastSentAt ?? "",
    fromEmail,
    source: existing?.source ?? "auto-cron",
    lastError,
    attempts: appendAttempt(existing, {
      at: attemptedAt,
      outcome: "error",
      message: lastError,
      fromEmail,
    }),
  };
  return upsertStatus(store, next, monthKey);
}
