import { currentMonthKey } from "@/lib/client-request-lists";

export type Bd15EmailAttachment = {
  id: string;
  clientId: string;
  /** YYYY-MM period this file belongs to. */
  monthKey: string;
  name: string;
  contentType: string;
  contentBytes: string;
  savedAt: string;
};

export type Bd15EmailAttachmentsStore = {
  attachments: Bd15EmailAttachment[];
};

export function createEmptyBd15EmailAttachmentsStore(): Bd15EmailAttachmentsStore {
  return { attachments: [] };
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

function newAttachmentId(): string {
  return `bd15-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function normalizeBd15EmailAttachment(
  value: unknown,
): Bd15EmailAttachment | null {
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
  const id = String(record.id ?? "").trim() || newAttachmentId();
  return {
    id,
    clientId,
    monthKey,
    name,
    contentType:
      String(record.contentType ?? "").trim() || "application/pdf",
    contentBytes,
    savedAt:
      typeof record.savedAt === "string" && record.savedAt.trim()
        ? record.savedAt.trim()
        : new Date().toISOString(),
  };
}

/** Drop attachments from prior months (new month = clear once). */
export function pruneBd15EmailAttachmentsToMonth(
  store: Bd15EmailAttachmentsStore,
  monthKey: string = currentMonthKey(),
): Bd15EmailAttachmentsStore {
  return {
    attachments: store.attachments.filter(
      (item) => item.monthKey === monthKey,
    ),
  };
}

export function normalizeBd15EmailAttachmentsStore(
  value: unknown,
  monthKey: string = currentMonthKey(),
): Bd15EmailAttachmentsStore {
  const record = asRecord(value);
  const list = Array.isArray(record?.attachments)
    ? record.attachments
    : Array.isArray(value)
      ? value
      : [];
  const attachments = list
    .map(normalizeBd15EmailAttachment)
    .filter((item): item is Bd15EmailAttachment => Boolean(item));
  return pruneBd15EmailAttachmentsToMonth({ attachments }, monthKey);
}

export function getBd15EmailAttachments(
  store: Bd15EmailAttachmentsStore,
  clientId: string,
  monthKey: string = currentMonthKey(),
): Bd15EmailAttachment[] {
  const id = clientId.trim();
  if (!id) {
    return [];
  }
  return store.attachments
    .filter((item) => item.clientId === id && item.monthKey === monthKey)
    .sort((a, b) => {
      const aTime = Date.parse(a.savedAt) || 0;
      const bTime = Date.parse(b.savedAt) || 0;
      return aTime - bTime;
    });
}

/** Append one file for the client/month (does not replace existing files). */
export function addBd15EmailAttachment(
  store: Bd15EmailAttachmentsStore,
  attachment: Omit<Bd15EmailAttachment, "id" | "savedAt"> & {
    id?: string;
    savedAt?: string;
  },
  options: { replaceSameName?: boolean } = {},
): Bd15EmailAttachmentsStore {
  const monthKey = currentMonthKey();
  const pruned = pruneBd15EmailAttachmentsToMonth(store, monthKey);
  const clientId = attachment.clientId.trim();
  const name = attachment.name.trim();
  const nameKey = name.toLowerCase();
  const withoutSameName = options.replaceSameName
    ? pruned.attachments.filter(
        (item) =>
          !(
            item.clientId === clientId &&
            item.monthKey === monthKey &&
            item.name.toLowerCase() === nameKey
          ),
      )
    : pruned.attachments;
  const next: Bd15EmailAttachment = {
    id: attachment.id?.trim() || newAttachmentId(),
    clientId,
    monthKey,
    name,
    contentType: attachment.contentType.trim() || "application/pdf",
    contentBytes: attachment.contentBytes,
    savedAt: new Date().toISOString(),
  };
  return { attachments: [...withoutSameName, next] };
}

export function removeBd15EmailAttachment(
  store: Bd15EmailAttachmentsStore,
  clientId: string,
  attachmentId: string,
  monthKey: string = currentMonthKey(),
): Bd15EmailAttachmentsStore {
  const id = clientId.trim();
  const fileId = attachmentId.trim();
  return {
    attachments: store.attachments.filter(
      (item) =>
        !(
          item.clientId === id &&
          item.monthKey === monthKey &&
          item.id === fileId
        ),
    ),
  };
}
