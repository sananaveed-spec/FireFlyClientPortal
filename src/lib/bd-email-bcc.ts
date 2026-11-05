export type BdEmailBccEntry = {
  clientId: string;
  scheduleId: string;
  /** Comma-separated BCC addresses. */
  bcc: string;
};

export type BdEmailBccStore = {
  entries: BdEmailBccEntry[];
};

export function createEmptyBdEmailBccStore(): BdEmailBccStore {
  return { entries: [] };
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

export function normalizeBdEmailBccEntry(
  value: unknown,
): BdEmailBccEntry | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const clientId = String(record.clientId ?? "").trim();
  const scheduleId = String(record.scheduleId ?? "").trim().toLowerCase();
  if (!clientId || !scheduleId) {
    return null;
  }
  return {
    clientId,
    scheduleId,
    bcc: String(record.bcc ?? "").trim(),
  };
}

export function normalizeBdEmailBccStore(value: unknown): BdEmailBccStore {
  const record = asRecord(value);
  const list = Array.isArray(record?.entries)
    ? record.entries
    : Array.isArray(value)
      ? value
      : [];
  const entries = list
    .map(normalizeBdEmailBccEntry)
    .filter((item): item is BdEmailBccEntry => Boolean(item));
  return { entries };
}

export function getBdEmailBcc(
  store: BdEmailBccStore,
  clientId: string,
  scheduleId: string,
): string {
  const id = clientId.trim();
  const schedule = scheduleId.trim().toLowerCase();
  if (!id || !schedule) {
    return "";
  }
  return (
    store.entries.find(
      (item) => item.clientId === id && item.scheduleId === schedule,
    )?.bcc ?? ""
  );
}

export function upsertBdEmailBcc(
  store: BdEmailBccStore,
  input: { clientId: string; scheduleId: string; bcc: string },
): BdEmailBccStore {
  const clientId = input.clientId.trim();
  const scheduleId = input.scheduleId.trim().toLowerCase();
  if (!clientId || !scheduleId) {
    return store;
  }
  const next: BdEmailBccEntry = {
    clientId,
    scheduleId,
    bcc: input.bcc.trim(),
  };
  const without = store.entries.filter(
    (item) =>
      !(item.clientId === clientId && item.scheduleId === scheduleId),
  );
  // Drop empty BCC rows so the file stays small.
  if (!next.bcc) {
    return { entries: without };
  }
  return { entries: [...without, next] };
}
