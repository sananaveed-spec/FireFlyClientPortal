import { currentMonthKey } from "@/lib/client-request-lists";

export type Bd15EmailStatusComment = {
  id: string;
  text: string;
  createdAt: string;
};

export type Bd15EmailStatus = {
  clientId: string;
  monthKey: string;
  emailSent: boolean;
  emailSentAt: string;
  comments: Bd15EmailStatusComment[];
};

export type Bd15EmailStatusStore = {
  statuses: Bd15EmailStatus[];
};

export function createEmptyBd15EmailStatusStore(): Bd15EmailStatusStore {
  return { statuses: [] };
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

function newId(prefix: string): string {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function normalizeBd15EmailStatusComment(
  value: unknown,
): Bd15EmailStatusComment | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const text = String(record.text ?? "").trim();
  if (!text) {
    return null;
  }
  return {
    id: String(record.id ?? "").trim() || newId("cmt"),
    text,
    createdAt:
      typeof record.createdAt === "string" && record.createdAt.trim()
        ? record.createdAt.trim()
        : new Date().toISOString(),
  };
}

export function normalizeBd15EmailStatus(
  value: unknown,
): Bd15EmailStatus | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const clientId = String(record.clientId ?? "").trim();
  const monthKey = String(record.monthKey ?? "").trim();
  if (!clientId || !monthKey) {
    return null;
  }
  const comments = Array.isArray(record.comments)
    ? record.comments
        .map(normalizeBd15EmailStatusComment)
        .filter((item): item is Bd15EmailStatusComment => Boolean(item))
        .sort((a, b) => {
          const aTime = Date.parse(a.createdAt) || 0;
          const bTime = Date.parse(b.createdAt) || 0;
          return bTime - aTime;
        })
    : [];
  return {
    clientId,
    monthKey,
    emailSent: Boolean(record.emailSent),
    emailSentAt:
      typeof record.emailSentAt === "string" && record.emailSentAt.trim()
        ? record.emailSentAt.trim()
        : "",
    comments,
  };
}

/** Drop status rows from prior months (new month = clear once). */
export function pruneBd15EmailStatusToMonth(
  store: Bd15EmailStatusStore,
  monthKey: string = currentMonthKey(),
): Bd15EmailStatusStore {
  return {
    statuses: store.statuses.filter((item) => item.monthKey === monthKey),
  };
}

export function normalizeBd15EmailStatusStore(
  value: unknown,
  monthKey: string = currentMonthKey(),
): Bd15EmailStatusStore {
  const record = asRecord(value);
  const list = Array.isArray(record?.statuses)
    ? record.statuses
    : Array.isArray(value)
      ? value
      : [];
  const statuses = list
    .map(normalizeBd15EmailStatus)
    .filter((item): item is Bd15EmailStatus => Boolean(item));
  return pruneBd15EmailStatusToMonth({ statuses }, monthKey);
}

export function getBd15EmailStatus(
  store: Bd15EmailStatusStore,
  clientId: string,
  monthKey: string = currentMonthKey(),
): Bd15EmailStatus {
  const id = clientId.trim();
  const existing = store.statuses.find(
    (item) => item.clientId === id && item.monthKey === monthKey,
  );
  if (existing) {
    return existing;
  }
  return {
    clientId: id,
    monthKey,
    emailSent: false,
    emailSentAt: "",
    comments: [],
  };
}

export function markBd15EmailSent(
  store: Bd15EmailStatusStore,
  clientId: string,
  monthKey: string = currentMonthKey(),
  sentAt: string = new Date().toISOString(),
): Bd15EmailStatusStore {
  const pruned = pruneBd15EmailStatusToMonth(store, monthKey);
  const current = getBd15EmailStatus(pruned, clientId, monthKey);
  const next: Bd15EmailStatus = {
    ...current,
    emailSent: true,
    emailSentAt: sentAt,
  };
  const without = pruned.statuses.filter(
    (item) =>
      !(item.clientId === next.clientId && item.monthKey === next.monthKey),
  );
  return { statuses: [...without, next] };
}

export function addBd15EmailStatusComment(
  store: Bd15EmailStatusStore,
  clientId: string,
  text: string,
  monthKey: string = currentMonthKey(),
): Bd15EmailStatusStore {
  const trimmed = text.trim();
  if (!trimmed) {
    return store;
  }
  const pruned = pruneBd15EmailStatusToMonth(store, monthKey);
  const current = getBd15EmailStatus(pruned, clientId, monthKey);
  const comment: Bd15EmailStatusComment = {
    id: newId("cmt"),
    text: trimmed,
    createdAt: new Date().toISOString(),
  };
  const next: Bd15EmailStatus = {
    ...current,
    comments: [comment, ...current.comments],
  };
  const without = pruned.statuses.filter(
    (item) =>
      !(item.clientId === next.clientId && item.monthKey === next.monthKey),
  );
  return { statuses: [...without, next] };
}
