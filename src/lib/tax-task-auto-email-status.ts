/**
 * Sent state for Annual / Quarterly due-date tax emails
 * (Corporate / Personal income tax, GST/HST return).
 */

export type TaxTaskAutoEmailKind =
  | "corporate-income-tax"
  | "personal-income-tax"
  | "gst-hst-return";

export type TaxTaskAutoEmailStatus = {
  profileId: string;
  kind: TaxTaskAutoEmailKind;
  /** Task due date YYYY-MM-DD that triggered the send. */
  dueDate: string;
  emailSent: boolean;
  lastSentAt: string;
  fromEmail: string;
  lastError: string;
};

export type TaxTaskAutoEmailStatusStore = {
  statuses: TaxTaskAutoEmailStatus[];
};

export function createEmptyTaxTaskAutoEmailStatusStore(): TaxTaskAutoEmailStatusStore {
  return { statuses: [] };
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

export function normalizeTaxTaskAutoEmailStatus(
  value: unknown,
): TaxTaskAutoEmailStatus | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const profileId = String(record.profileId ?? "").trim();
  const kind = String(record.kind ?? "").trim() as TaxTaskAutoEmailKind;
  const dueDate = String(record.dueDate ?? "").trim();
  if (
    !profileId ||
    !dueDate ||
    (kind !== "corporate-income-tax" &&
      kind !== "personal-income-tax" &&
      kind !== "gst-hst-return")
  ) {
    return null;
  }
  return {
    profileId,
    kind,
    dueDate,
    emailSent: Boolean(record.emailSent),
    lastSentAt: String(record.lastSentAt ?? "").trim(),
    fromEmail: String(record.fromEmail ?? "").trim(),
    lastError: String(record.lastError ?? "").trim(),
  };
}

export function normalizeTaxTaskAutoEmailStatusStore(
  value: unknown,
): TaxTaskAutoEmailStatusStore {
  const record = asRecord(value);
  const list = Array.isArray(record?.statuses) ? record!.statuses : [];
  return {
    statuses: list
      .map(normalizeTaxTaskAutoEmailStatus)
      .filter((item): item is TaxTaskAutoEmailStatus => Boolean(item)),
  };
}

export function getTaxTaskAutoEmailStatus(
  store: TaxTaskAutoEmailStatusStore,
  profileId: string,
  kind: TaxTaskAutoEmailKind,
  dueDate: string,
): TaxTaskAutoEmailStatus | null {
  const id = profileId.trim();
  const due = dueDate.trim();
  return (
    store.statuses.find(
      (item) =>
        item.profileId === id && item.kind === kind && item.dueDate === due,
    ) ?? null
  );
}

export function markTaxTaskAutoEmailSent(
  store: TaxTaskAutoEmailStatusStore,
  input: {
    profileId: string;
    kind: TaxTaskAutoEmailKind;
    dueDate: string;
    lastSentAt: string;
    fromEmail: string;
  },
): TaxTaskAutoEmailStatusStore {
  const profileId = input.profileId.trim();
  const dueDate = input.dueDate.trim();
  const next: TaxTaskAutoEmailStatus = {
    profileId,
    kind: input.kind,
    dueDate,
    emailSent: true,
    lastSentAt: input.lastSentAt,
    fromEmail: input.fromEmail.trim(),
    lastError: "",
  };
  const statuses = store.statuses.filter(
    (item) =>
      !(
        item.profileId === profileId &&
        item.kind === input.kind &&
        item.dueDate === dueDate
      ),
  );
  return { statuses: [...statuses, next] };
}

export function markTaxTaskAutoEmailError(
  store: TaxTaskAutoEmailStatusStore,
  input: {
    profileId: string;
    kind: TaxTaskAutoEmailKind;
    dueDate: string;
    lastError: string;
    fromEmail: string;
  },
): TaxTaskAutoEmailStatusStore {
  const profileId = input.profileId.trim();
  const dueDate = input.dueDate.trim();
  const prior = getTaxTaskAutoEmailStatus(store, profileId, input.kind, dueDate);
  const next: TaxTaskAutoEmailStatus = {
    profileId,
    kind: input.kind,
    dueDate,
    emailSent: prior?.emailSent ?? false,
    lastSentAt: prior?.lastSentAt ?? "",
    fromEmail: input.fromEmail.trim() || prior?.fromEmail || "",
    lastError: input.lastError.trim(),
  };
  const statuses = store.statuses.filter(
    (item) =>
      !(
        item.profileId === profileId &&
        item.kind === input.kind &&
        item.dueDate === dueDate
      ),
  );
  return { statuses: [...statuses, next] };
}
