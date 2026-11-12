/**
 * Client Info — One-Time Tasks checklist.
 * Ignition "Bookkeeping Services" heading seeds QBO Setup.
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

/** Default seeded tasks (e.g. QBO Setup) cannot be edited or deleted. */
export function isProtectedOneTimeTask(task: Pick<OneTimeTaskItem, "kind" | "source">): boolean {
  return (
    task.kind === QBO_SETUP_TASK_KIND || task.source === "ignition_bookkeeping"
  );
}

export function isManualOneTimeTask(task: Pick<OneTimeTaskItem, "kind" | "source">): boolean {
  return !isProtectedOneTimeTask(task);
}

export function updateOneTimeTaskLabel(
  tasks: OneTimeTaskItem[],
  taskId: string,
  label: string,
): OneTimeTaskItem[] {
  const nextLabel = label.trim();
  if (!nextLabel) {
    return tasks;
  }
  return tasks.map((task) => {
    if (task.id !== taskId || isProtectedOneTimeTask(task)) {
      return task;
    }
    return { ...task, label: nextLabel };
  });
}

export function removeOneTimeTask(
  tasks: OneTimeTaskItem[],
  taskId: string,
): OneTimeTaskItem[] {
  return tasks.filter((task) => {
    if (task.id !== taskId) {
      return true;
    }
    // Never drop protected default tasks.
    return isProtectedOneTimeTask(task);
  });
}

/**
 * When saving, keep any protected defaults that were omitted from the payload.
 */
export function mergeProtectedOneTimeTasks(
  previous: OneTimeTaskItem[],
  next: OneTimeTaskItem[],
): OneTimeTaskItem[] {
  const protectedPrevious = previous.filter(isProtectedOneTimeTask);
  const nextIds = new Set(next.map((task) => task.id));
  const missing = protectedPrevious.filter((task) => !nextIds.has(task.id));
  if (missing.length === 0) {
    return next.map((task) => {
      if (!isProtectedOneTimeTask(task)) {
        return task;
      }
      const prior = previous.find((item) => item.id === task.id);
      if (!prior) {
        return {
          ...task,
          kind: QBO_SETUP_TASK_KIND,
          label: QBO_SETUP_TASK_LABEL,
          source: "ignition_bookkeeping" as const,
        };
      }
      return {
        ...prior,
        completed: task.completed,
        completedAt: task.completedAt,
        completedByName: task.completedByName,
        completedByEmail: task.completedByEmail,
      };
    });
  }
  return [...missing, ...next];
}

/**
 * True when an Ignition proposal has a service-group heading
 * "Bookkeeping Services" or "Bookkeeping Service".
 */
export function proposalHasBookkeepingServicesHeading(
  services: ReadonlyArray<{ groupName?: string; name?: string }>,
): boolean {
  return services.some((service) => {
    const heading = String(service.groupName ?? "")
      .trim()
      .replace(/\s+/g, " ");
    return /^bookkeeping\s+services?$/i.test(heading);
  });
}

/** @deprecated Use proposalHasBookkeepingServicesHeading */
export function proposalServicesIncludeBookkeeping(
  services: ReadonlyArray<{ groupName?: string; name?: string }>,
): boolean {
  return proposalHasBookkeepingServicesHeading(services);
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
 * If the proposal has a Bookkeeping Services heading and QBO Setup is not
 * already listed, append the single auto checkbox (no duplicates).
 */
export function ensureQboSetupTaskFromBookkeeping(
  tasks: OneTimeTaskItem[],
  hasBookkeepingHeading: boolean,
  proposalSlug = "",
): OneTimeTaskItem[] {
  if (!hasBookkeepingHeading || hasQboSetupOneTimeTask(tasks)) {
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
