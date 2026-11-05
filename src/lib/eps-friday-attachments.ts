import { currentMonthKey } from "@/lib/client-request-lists";
import { isEpsFridayWeekScheduleId } from "@/lib/eps-friday-weeks";

export type EpsFridayAttachment = {
  id: string;
  clientId: string;
  /** e.g. eps-week-1-friday */
  scheduleId: string;
  monthKey: string;
  name: string;
  contentType: string;
  contentBytes: string;
  savedAt: string;
};

export type EpsFridayAttachmentsStore = {
  attachments: EpsFridayAttachment[];
};

export function createEmptyEpsFridayAttachmentsStore(): EpsFridayAttachmentsStore {
  return { attachments: [] };
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

function newAttachmentId(): string {
  return `eps-fri-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function isAllowedEpsFridaySpreadsheetName(name: string): boolean {
  const lower = name.trim().toLowerCase();
  return (
    lower.endsWith(".xlsx") ||
    lower.endsWith(".xls") ||
    lower.endsWith(".csv")
  );
}

export function normalizeEpsFridayAttachment(
  value: unknown,
): EpsFridayAttachment | null {
  const record = asRecord(value);
  if (!record) {
    return null;
  }
  const clientId = String(record.clientId ?? "").trim();
  const scheduleId = String(record.scheduleId ?? "").trim().toLowerCase();
  const monthKey = String(record.monthKey ?? "").trim();
  const name = String(record.name ?? "").trim();
  const contentBytes = String(record.contentBytes ?? "").trim();
  if (
    !clientId ||
    !monthKey ||
    !name ||
    !contentBytes ||
    !isEpsFridayWeekScheduleId(scheduleId)
  ) {
    return null;
  }
  return {
    id: String(record.id ?? "").trim() || newAttachmentId(),
    clientId,
    scheduleId,
    monthKey,
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

export function pruneEpsFridayAttachmentsToMonth(
  store: EpsFridayAttachmentsStore,
  monthKey: string = currentMonthKey(),
): EpsFridayAttachmentsStore {
  return {
    attachments: store.attachments.filter(
      (item) => item.monthKey === monthKey,
    ),
  };
}

export function normalizeEpsFridayAttachmentsStore(
  value: unknown,
  monthKey: string = currentMonthKey(),
): EpsFridayAttachmentsStore {
  const record = asRecord(value);
  const list = Array.isArray(record?.attachments)
    ? record.attachments
    : Array.isArray(value)
      ? value
      : [];
  const attachments = list
    .map(normalizeEpsFridayAttachment)
    .filter((item): item is EpsFridayAttachment => Boolean(item));
  return pruneEpsFridayAttachmentsToMonth({ attachments }, monthKey);
}

export function getEpsFridayAttachments(
  store: EpsFridayAttachmentsStore,
  clientId: string,
  scheduleId: string,
  monthKey: string = currentMonthKey(),
): EpsFridayAttachment[] {
  const id = clientId.trim();
  const schedule = scheduleId.trim().toLowerCase();
  if (!id || !isEpsFridayWeekScheduleId(schedule)) {
    return [];
  }
  return store.attachments
    .filter(
      (item) =>
        item.clientId === id &&
        item.scheduleId === schedule &&
        item.monthKey === monthKey,
    )
    .sort((a, b) => {
      const aTime = Date.parse(a.savedAt) || 0;
      const bTime = Date.parse(b.savedAt) || 0;
      return aTime - bTime;
    });
}

export function addEpsFridayAttachment(
  store: EpsFridayAttachmentsStore,
  attachment: Omit<EpsFridayAttachment, "id" | "savedAt"> & {
    id?: string;
    savedAt?: string;
  },
): EpsFridayAttachmentsStore {
  const monthKey = attachment.monthKey.trim() || currentMonthKey();
  const pruned = pruneEpsFridayAttachmentsToMonth(store, monthKey);
  const clientId = attachment.clientId.trim();
  const scheduleId = attachment.scheduleId.trim().toLowerCase();
  const name = attachment.name.trim();
  const nameKey = name.toLowerCase();
  const withoutSameName = pruned.attachments.filter(
    (item) =>
      !(
        item.clientId === clientId &&
        item.scheduleId === scheduleId &&
        item.monthKey === monthKey &&
        item.name.toLowerCase() === nameKey
      ),
  );
  const next: EpsFridayAttachment = {
    id: attachment.id?.trim() || newAttachmentId(),
    clientId,
    scheduleId,
    monthKey,
    name,
    contentType: attachment.contentType.trim() || "application/octet-stream",
    contentBytes: attachment.contentBytes,
    savedAt: new Date().toISOString(),
  };
  return { attachments: [...withoutSameName, next] };
}

export function removeEpsFridayAttachment(
  store: EpsFridayAttachmentsStore,
  clientId: string,
  scheduleId: string,
  attachmentId: string,
  monthKey: string = currentMonthKey(),
): EpsFridayAttachmentsStore {
  const id = clientId.trim();
  const schedule = scheduleId.trim().toLowerCase();
  const fileId = attachmentId.trim();
  return {
    attachments: store.attachments.filter(
      (item) =>
        !(
          item.clientId === id &&
          item.scheduleId === schedule &&
          item.monthKey === monthKey &&
          item.id === fileId
        ),
    ),
  };
}

export function clearEpsFridayAttachments(
  store: EpsFridayAttachmentsStore,
  clientId: string,
  scheduleId: string,
  monthKey: string = currentMonthKey(),
): EpsFridayAttachmentsStore {
  const id = clientId.trim();
  const schedule = scheduleId.trim().toLowerCase();
  return {
    attachments: store.attachments.filter(
      (item) =>
        !(
          item.clientId === id &&
          item.scheduleId === schedule &&
          item.monthKey === monthKey
        ),
    ),
  };
}
