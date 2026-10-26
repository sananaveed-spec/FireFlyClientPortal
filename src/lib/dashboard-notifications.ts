import { copyFile, mkdir, readFile, unlink, writeFile } from "fs/promises";
import path from "path";
import { getDataDir } from "@/lib/data-store";

export type DashboardNotificationKind = "ignition_new_client";

export type DashboardNotification = {
  id: string;
  kind: DashboardNotificationKind;
  title: string;
  message: string;
  clientName: string;
  clientProfileId: string;
  createdAt: string;
  /** ISO when dismissed; empty = still active. */
  dismissedAt: string;
};

export type DashboardNotificationsStore = {
  notifications: DashboardNotification[];
};

function notificationsPath() {
  return path.join(getDataDir(), "dashboard-notifications.json");
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

export function createEmptyDashboardNotificationsStore(): DashboardNotificationsStore {
  return { notifications: [] };
}

export function createDashboardNotificationId() {
  return `note-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function normalizeNotification(value: unknown): DashboardNotification | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<DashboardNotification>;
  const id = String(record.id ?? "").trim();
  const title = String(record.title ?? "").trim();
  const message = String(record.message ?? "").trim();
  if (!id || (!title && !message)) {
    return null;
  }
  return {
    id,
    kind:
      record.kind === "ignition_new_client"
        ? "ignition_new_client"
        : "ignition_new_client",
    title,
    message,
    clientName: String(record.clientName ?? "").trim(),
    clientProfileId: String(record.clientProfileId ?? "").trim(),
    createdAt:
      String(record.createdAt ?? "").trim() || new Date().toISOString(),
    dismissedAt: String(record.dismissedAt ?? "").trim(),
  };
}

export function normalizeDashboardNotificationsStore(
  value: unknown,
): DashboardNotificationsStore {
  const record =
    typeof value === "object" && value !== null
      ? (value as { notifications?: unknown })
      : {};
  const notifications = Array.isArray(record.notifications)
    ? record.notifications
        .map(normalizeNotification)
        .filter((item): item is DashboardNotification => Boolean(item))
    : [];
  return { notifications };
}

export async function readDashboardNotificationsStore(): Promise<DashboardNotificationsStore> {
  await ensureDataDir();
  try {
    const raw = await readFile(notificationsPath(), "utf8");
    return normalizeDashboardNotificationsStore(
      JSON.parse(raw.replace(/^\uFEFF/, "")),
    );
  } catch (error) {
    if (isMissingFile(error)) {
      return createEmptyDashboardNotificationsStore();
    }
    throw error;
  }
}

export async function writeDashboardNotificationsStore(
  store: DashboardNotificationsStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(notificationsPath(), store);
}

export function listActiveDashboardNotifications(
  store: DashboardNotificationsStore,
): DashboardNotification[] {
  return store.notifications
    .filter((item) => !item.dismissedAt)
    .sort(
      (a, b) =>
        new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
    );
}

export async function appendDashboardNotifications(
  entries: Array<Omit<DashboardNotification, "id" | "createdAt" | "dismissedAt"> & {
    createdAt?: string;
  }>,
): Promise<DashboardNotification[]> {
  if (entries.length === 0) {
    return [];
  }
  const store = await readDashboardNotificationsStore();
  const created: DashboardNotification[] = entries.map((entry) => ({
    id: createDashboardNotificationId(),
    kind: entry.kind,
    title: entry.title,
    message: entry.message,
    clientName: entry.clientName,
    clientProfileId: entry.clientProfileId,
    createdAt: entry.createdAt ?? new Date().toISOString(),
    dismissedAt: "",
  }));
  const next: DashboardNotificationsStore = {
    notifications: [...created, ...store.notifications].slice(0, 200),
  };
  await writeDashboardNotificationsStore(next);
  return created;
}

export async function dismissDashboardNotification(
  id: string,
): Promise<DashboardNotificationsStore> {
  const store = await readDashboardNotificationsStore();
  const now = new Date().toISOString();
  const next: DashboardNotificationsStore = {
    notifications: store.notifications.map((item) =>
      item.id === id && !item.dismissedAt
        ? { ...item, dismissedAt: now }
        : item,
    ),
  };
  await writeDashboardNotificationsStore(next);
  return next;
}
