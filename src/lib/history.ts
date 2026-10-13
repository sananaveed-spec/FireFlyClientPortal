export type HistoryEntityType =
  | "clientProfile"
  | "resourceAllocation"
  | "lookupData";

export type HistoryAction = "created" | "updated" | "deleted";

export type HistoryEntry = {
  id: string;
  /** ISO timestamp */
  createdAt: string;
  employeeName: string;
  employeeEmail: string;
  /** Client / project name related to the action */
  project: string;
  entityType: HistoryEntityType;
  entityLabel: string;
  action: HistoryAction;
  summary: string;
  details: string;
  comment: string;
};

export type HistoryInput = Omit<HistoryEntry, "id" | "createdAt"> & {
  createdAt?: string;
};

export function createHistoryId() {
  return `hist-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function formatHistoryDate(iso: string) {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return "—";
  }
  return date.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

export function formatHistoryTime(iso: string) {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return "—";
  }
  return date.toLocaleTimeString(undefined, {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

export function entityTypeLabel(type: HistoryEntityType) {
  switch (type) {
    case "clientProfile":
      return "Client Profile";
    case "resourceAllocation":
      return "Resource Allocation";
    case "lookupData":
      return "Lookup Data";
    default:
      return type;
  }
}

export function actionLabel(action: HistoryAction) {
  switch (action) {
    case "created":
      return "Created";
    case "updated":
      return "Updated";
    case "deleted":
      return "Deleted";
    default:
      return action;
  }
}

export function normalizeHistoryEntry(
  value: Partial<HistoryEntry> & { id?: string },
): HistoryEntry | null {
  const employeeName = String(value.employeeName ?? "").trim();
  const summary = String(value.summary ?? "").trim();
  if (!employeeName && !summary) {
    return null;
  }

  const entityType =
    value.entityType === "clientProfile" ||
    value.entityType === "resourceAllocation" ||
    value.entityType === "lookupData"
      ? value.entityType
      : "resourceAllocation";

  const action =
    value.action === "created" ||
    value.action === "updated" ||
    value.action === "deleted"
      ? value.action
      : "updated";

  return {
    id:
      typeof value.id === "string" && value.id
        ? value.id
        : createHistoryId(),
    createdAt:
      typeof value.createdAt === "string" && value.createdAt
        ? value.createdAt
        : new Date().toISOString(),
    employeeName: employeeName || "Unknown",
    employeeEmail: String(value.employeeEmail ?? "").trim(),
    project: String(value.project ?? "").trim(),
    entityType,
    entityLabel: String(value.entityLabel ?? "").trim(),
    action,
    summary: summary || actionLabel(action),
    details: String(value.details ?? "").trim(),
    comment: String(value.comment ?? "").trim(),
  };
}

/** Build a short before→after style details string for changed fields. */
export function buildChangeDetails(
  before: Record<string, string>,
  after: Record<string, string>,
): string {
  const keys = [...new Set([...Object.keys(before), ...Object.keys(after)])];
  const lines: string[] = [];
  for (const key of keys) {
    const prev = (before[key] ?? "").trim();
    const next = (after[key] ?? "").trim();
    if (prev === next) {
      continue;
    }
    lines.push(`${key}: "${prev || "—"}" → "${next || "—"}"`);
  }
  return lines.join("\n");
}

export function snapshotFromLabeledFields(
  fields: { label: string; value: string }[],
): Record<string, string> {
  const snapshot: Record<string, string> = {};
  for (const field of fields) {
    snapshot[field.label] = field.value;
  }
  return snapshot;
}
