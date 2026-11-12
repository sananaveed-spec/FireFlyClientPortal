/**
 * Client Info — Quarterly Tasks checklist.
 */

export type QuarterlyTaskItem = {
  id: string;
  /** Reserved for future default/auto tasks; empty for user-added. */
  kind: string;
  label: string;
  completed: boolean;
  completedAt: string;
  completedByName: string;
  completedByEmail: string;
  createdAt: string;
  source: "manual";
  proposalSlug: string;
};

export function createQuarterlyTaskId() {
  return `qot-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function normalizeQuarterlyTaskItem(
  value: unknown,
): QuarterlyTaskItem | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<QuarterlyTaskItem>;
  const id = String(record.id ?? "").trim();
  const label = String(record.label ?? "").trim();
  if (!id || !label) {
    return null;
  }
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
    source: "manual",
    proposalSlug: String(record.proposalSlug ?? "").trim(),
  };
}

export function normalizeQuarterlyTasks(value: unknown): QuarterlyTaskItem[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map(normalizeQuarterlyTaskItem)
    .filter((item): item is QuarterlyTaskItem => Boolean(item));
}

export function quarterlyTasksStatusLabel(tasks: QuarterlyTaskItem[]): string {
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

export function isProtectedQuarterlyTask(
  task: Pick<QuarterlyTaskItem, "kind" | "source">,
): boolean {
  return Boolean(String(task.kind ?? "").trim());
}

export function isManualQuarterlyTask(
  task: Pick<QuarterlyTaskItem, "kind" | "source">,
): boolean {
  return !isProtectedQuarterlyTask(task);
}

export function createManualQuarterlyTask(label: string): QuarterlyTaskItem {
  return {
    id: createQuarterlyTaskId(),
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

export function toggleQuarterlyTaskCompletion(
  tasks: QuarterlyTaskItem[],
  taskId: string,
  completed: boolean,
  actor: { name: string; email: string },
): QuarterlyTaskItem[] {
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

export function updateQuarterlyTaskLabel(
  tasks: QuarterlyTaskItem[],
  taskId: string,
  label: string,
): QuarterlyTaskItem[] {
  const nextLabel = label.trim();
  return tasks.map((task) => {
    if (task.id !== taskId || isProtectedQuarterlyTask(task) || !nextLabel) {
      return task;
    }
    return { ...task, label: nextLabel };
  });
}

export function removeQuarterlyTask(
  tasks: QuarterlyTaskItem[],
  taskId: string,
): QuarterlyTaskItem[] {
  return tasks.filter((task) => {
    if (task.id !== taskId) {
      return true;
    }
    return isProtectedQuarterlyTask(task);
  });
}

export function mergeProtectedQuarterlyTasks(
  previous: QuarterlyTaskItem[],
  next: QuarterlyTaskItem[],
): QuarterlyTaskItem[] {
  const protectedPrevious = previous.filter(isProtectedQuarterlyTask);
  const nextIds = new Set(next.map((task) => task.id));
  const nextKinds = new Set(
    next.filter(isProtectedQuarterlyTask).map((task) => task.kind),
  );
  const missing = protectedPrevious.filter(
    (task) => !nextIds.has(task.id) && !nextKinds.has(task.kind),
  );
  const merged = next.map((task) => {
    if (!isProtectedQuarterlyTask(task)) {
      return task;
    }
    const prior =
      previous.find((item) => item.id === task.id) ||
      previous.find((item) => item.kind && item.kind === task.kind);
    if (!prior) {
      return task;
    }
    return {
      ...prior,
      label: prior.label,
      kind: prior.kind,
      source: "manual" as const,
      completed: task.completed,
      completedAt: task.completedAt,
      completedByName: task.completedByName,
      completedByEmail: task.completedByEmail,
    };
  });
  return missing.length > 0 ? [...missing, ...merged] : merged;
}
