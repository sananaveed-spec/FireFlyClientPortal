/**
 * Client Info — One-Time Tasks checklist (Ignition bookkeeping seeds QBO Setup).
 */

export const QBO_SETUP_TASK_KIND = "qbo_setup";

export const QBO_SETUP_TASK_LABEL =
  "QuickBooks Online (QBO) Setup: Setting up QBO, including creating the chart of accounts, adding products and services, entering customer and vendor details, and connecting bank and credit card accounts.";

export type OneTimeTaskItem = {
  id: string;
  /** Stable kind for auto-seeded tasks; empty for user-added. */
  kind: string;
  label: string;
  completed: boolean;
  completedAt: string;
  completedByName: string;
  completedByEmail: string;
  createdAt: string;
  source: "ignition_bookkeeping" | "manual";
  proposalSlug: string;
};

export function createOneTimeTaskId() {
  return `ott-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function normalizeOneTimeTaskItem(value: unknown): OneTimeTaskItem | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<OneTimeTaskItem>;
  const id = String(record.id ?? "").trim();
  const label = String(record.label ?? "").trim();
  if (!id || !label) {
    return null;
  }
  const source =
    record.source === "ignition_bookkeeping" ? "ignition_bookkeeping" : "manual";
  return {
    id,
    kind: String(record.kind ?? "").trim(),
    label,
    completed: Boolean(record.completed),
    completedAt: String(record.completedAt ?? "").trim(),
    completedByName: String(record.completedByName ?? "").trim(),
    completedByEmail: String(record.completedByEmail ?? "").trim(),
    createdAt:
      String(record.createdAt ?? "").trim() || new Date().toISOString(),
    source,
    proposalSlug: String(record.proposalSlug ?? "").trim(),
  };
}

export function normalizeOneTimeTasks(value: unknown): OneTimeTaskItem[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map(normalizeOneTimeTaskItem)
    .filter((item): item is OneTimeTaskItem => Boolean(item));
}

export function oneTimeTasksStatusLabel(tasks: OneTimeTaskItem[]): string {
  if (tasks.length === 0) {
    return "";
  }
  if (tasks.every((task) => task.completed)) {
    return "Completed";
  }
  if (tasks.some((task) => task.completed)) {
    return "In progress";
  }
  return "Pending";
}

export function createManualOneTimeTask(label: string): OneTimeTaskItem {
  return {
    id: createOneTimeTaskId(),
    kind: "",
    label: label.trim(),
    completed: false,
    completedAt: "",
    completedByName: "",
    completedByEmail: "",
    createdAt: new Date().toISOString(),
    source: "manual",
    proposalSlug: "",
  };
}

export function createQboSetupOneTimeTask(proposalSlug = ""): OneTimeTaskItem {
  return {
    id: createOneTimeTaskId(),
    kind: QBO_SETUP_TASK_KIND,
    label: QBO_SETUP_TASK_LABEL,
    completed: false,
    completedAt: "",
    completedByName: "",
    completedByEmail: "",
    createdAt: new Date().toISOString(),
    source: "ignition_bookkeeping",
    proposalSlug: proposalSlug.trim(),
  };
}

export function hasQboSetupOneTimeTask(tasks: readonly OneTimeTaskItem[]): boolean {
  return tasks.some((task) => task.kind === QBO_SETUP_TASK_KIND);
}

/** True when Ignition proposal service names look like bookkeeping. */
export function proposalServicesIncludeBookkeeping(
  services: ReadonlyArray<{ name?: string }>,
): boolean {
  return services.some((service) =>
    /bookkeep/i.test(String(service.name ?? "").trim()),
  );
}

/**
 * True when Resource Allocation has Bookkeeping services selected.
 * Any value in that multi-select counts (Monthly / Weekly / etc.).
 */
export function resourceServicesIncludeBookkeeping(
  bookkeeping: ReadonlyArray<string> | undefined | null,
): boolean {
  return (
    Array.isArray(bookkeeping) &&
    bookkeeping.some((item) => String(item ?? "").trim().length > 0)
  );
}

/**
 * If bookkeeping is on the proposal and QBO Setup is not already listed,
 * append the single auto checkbox (does not add duplicates).
 */
export function ensureQboSetupTaskFromBookkeeping(
  tasks: OneTimeTaskItem[],
  hasBookkeeping: boolean,
  proposalSlug = "",
): OneTimeTaskItem[] {
  if (!hasBookkeeping || hasQboSetupOneTimeTask(tasks)) {
    return tasks;
  }
  return [...tasks, createQboSetupOneTimeTask(proposalSlug)];
}

export function toggleOneTimeTaskCompletion(
  tasks: OneTimeTaskItem[],
  taskId: string,
  completed: boolean,
  actor: { name: string; email: string },
): OneTimeTaskItem[] {
  const now = new Date().toISOString();
  return tasks.map((task) => {
    if (task.id !== taskId) {
      return task;
    }
    if (completed) {
      return {
        ...task,
        completed: true,
        completedAt: now,
        completedByName: actor.name.trim() || actor.email.trim(),
        completedByEmail: actor.email.trim(),
      };
    }
    return {
      ...task,
      completed: false,
      completedAt: "",
      completedByName: "",
      completedByEmail: "",
    };
  });
}
