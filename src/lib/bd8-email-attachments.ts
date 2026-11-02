import { currentMonthKey } from "@/lib/client-request-lists";

export type BdEmailAttachmentScheduleId = "bd8" | "bd13";

export type Bd8EmailAttachment = {
  clientId: string;
  /** YYYY-MM period this file belongs to. */
  monthKey: string;
  /** Which schedule owns this file (default bd8 for legacy rows). */
  scheduleId: BdEmailAttachmentScheduleId;
  name: string;
  contentType: string;
  contentBytes: string;
  savedAt: string;
};

export type Bd8EmailAttachmentsStore = {
  attachments: Bd8EmailAttachment[];
};

export function createEmptyBd8EmailAttachmentsStore(): Bd8EmailAttachmentsStore {
  return { attachments: [] };
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

export function normalizeBdEmailAttachmentScheduleId(
  value: unknown,
): BdEmailAttachmentScheduleId {
  return String(value ?? "").trim().toLowerCase() === "bd13" ? "bd13" : "bd8";
}

export function normalizeBd8EmailAttachment(
  value: unknown,
): Bd8EmailAttachment | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const clientId = String(record.clientId ?? "").trim();
  const monthKey = String(record.monthKey ?? "").trim();
  const name = String(record.name ?? "").trim();
  const contentBytes = String(record.contentBytes ?? "").trim();
  if (!clientId || !monthKey || !name || !contentBytes) {
    return null;
  }
  return {
    clientId,
    monthKey,
    scheduleId: normalizeBdEmailAttachmentScheduleId(record.scheduleId),
    name,
    contentType:
      String(record.contentType ?? "").trim() || "application/octet-stream",
    contentBytes,
    savedAt:
      typeof record.savedAt === "string" && record.savedAt.trim()
        ? record.savedAt.trim()
        : new Date().toISOString(),
  };
}

/** Drop attachments from prior months (new month = clear once). */
export function pruneBd8EmailAttachmentsToMonth(
  store: Bd8EmailAttachmentsStore,
  monthKey: string = currentMonthKey(),
): Bd8EmailAttachmentsStore {
  return {
    attachments: store.attachments.filter(
      (item) => item.monthKey === monthKey,
    ),
  };
}

export function normalizeBd8EmailAttachmentsStore(
  value: unknown,
  monthKey: string = currentMonthKey(),
): Bd8EmailAttachmentsStore {
  const record = asRecord(value);
  const list = Array.isArray(record?.attachments)
    ? record.attachments
    : Array.isArray(value)
      ? value
      : [];
  const attachments = list
    .map(normalizeBd8EmailAttachment)
    .filter((item): item is Bd8EmailAttachment => Boolean(item));
  return pruneBd8EmailAttachmentsToMonth({ attachments }, monthKey);
}

export function getBd8EmailAttachment(
  store: Bd8EmailAttachmentsStore,
  clientId: string,
  monthKey: string = currentMonthKey(),
  scheduleId: BdEmailAttachmentScheduleId = "bd8",
): Bd8EmailAttachment | null {
  const id = clientId.trim();
  if (!id) {
    return null;
  }
  return (
    store.attachments.find(
      (item) =>
        item.clientId === id &&
        item.monthKey === monthKey &&
        item.scheduleId === scheduleId,
    ) ?? null
  );
}

export function upsertBd8EmailAttachment(
  store: Bd8EmailAttachmentsStore,
  attachment: Bd8EmailAttachment,
): Bd8EmailAttachmentsStore {
  const monthKey = currentMonthKey();
  const pruned = pruneBd8EmailAttachmentsToMonth(store, monthKey);
  const next: Bd8EmailAttachment = {
    ...attachment,
    monthKey,
    scheduleId: normalizeBdEmailAttachmentScheduleId(attachment.scheduleId),
    savedAt: new Date().toISOString(),
  };
  const without = pruned.attachments.filter(
    (item) =>
      !(
        item.clientId === next.clientId &&
        item.monthKey === next.monthKey &&
        item.scheduleId === next.scheduleId
      ),
  );
  return { attachments: [...without, next] };
}

export function clearBd8EmailAttachment(
  store: Bd8EmailAttachmentsStore,
  clientId: string,
  monthKey: string = currentMonthKey(),
  scheduleId: BdEmailAttachmentScheduleId = "bd8",
): Bd8EmailAttachmentsStore {
  const id = clientId.trim();
  return {
    attachments: store.attachments.filter(
      (item) =>
        !(
          item.clientId === id &&
          item.monthKey === monthKey &&
          item.scheduleId === scheduleId
        ),
    ),
  };
}
