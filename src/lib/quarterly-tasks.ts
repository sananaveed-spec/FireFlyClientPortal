/**
 * Client Info — Quarterly Tasks checklist.
 * GST/HST Tax Compliance Service(s) proposal heading seeds the GST/HST return
 * bullet; due dates are set by assigned users.
 */

export const QUARTERLY_GST_HST_TASK_SOURCE = "ignition_gst_hst" as const;

/** Fixed Quarterly Task for GST/HST Tax Compliance Service(s) heading. */
export const GST_HST_TAX_COMPLIANCE_QUARTERLY_TASK_DEFINITIONS = [
  {
    kind: "gst_hst_tax_compliance:return-filing",
    label: "Monthly / quarterly GST/HST return preparation and filing.",
  },
] as const;

export type QuarterlyTaskItem = {
  id: string;
  /** Stable key for default tasks; empty for user-added. */
  kind: string;
  label: string;
  /** ISO date YYYY-MM-DD when set by an assigned user. */
  dueDate: string;
  completed: boolean;
  completedAt: string;
  completedByName: string;
  completedByEmail: string;
  createdAt: string;
  source: typeof QUARTERLY_GST_HST_TASK_SOURCE | "manual";
  proposalSlug: string;
};

export function createQuarterlyTaskId() {
  return `qot-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function formatQuarterlyDueDateLabel(dueDate: string): string {
  const match = dueDate.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return dueDate.trim() || "—";
  }
  const date = new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
  );
  if (Number.isNaN(date.getTime())) {
    return dueDate;
  }
  return date.toLocaleDateString(undefined, {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function gstHstTaxComplianceKindForLabel(label: string): string {
  const normalized = label.trim().toLowerCase().replace(/\s+/g, " ");
  const match = GST_HST_TAX_COMPLIANCE_QUARTERLY_TASK_DEFINITIONS.find(
    (def) => def.label.toLowerCase() === normalized,
  );
  return match?.kind ?? "";
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
  const kindFromLabel = gstHstTaxComplianceKindForLabel(label);
  const kind = String(record.kind ?? "").trim() || kindFromLabel;
  return {
    id,
    kind,
    label,
    dueDate: String(record.dueDate ?? "").trim(),
    completed: Boolean(record.completed),
    completedAt: String(record.completedAt ?? "").trim(),
    completedByName: String(record.completedByName ?? "").trim(),
    completedByEmail: String(record.completedByEmail ?? "").trim(),
    createdAt:
      String(record.createdAt ?? "").trim() || new Date().toISOString(),
    source:
      kind.startsWith("gst_hst_tax_compliance:") ||
      record.source === QUARTERLY_GST_HST_TASK_SOURCE
        ? QUARTERLY_GST_HST_TASK_SOURCE
        : "manual",
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
  const kind = String(task.kind ?? "");
  return (
    task.source === QUARTERLY_GST_HST_TASK_SOURCE ||
    kind.startsWith("gst_hst_tax_compliance:")
  );
}

/** Lookup template label for due-date auto-send notice (Template 25). */
export function quarterlyTaskAutoEmailTemplateName(
  task: Pick<QuarterlyTaskItem, "kind">,
): string {
  const kind = String(task.kind ?? "");
  if (kind === "gst_hst_tax_compliance:return-filing") {
    return "Template 25";
  }
  return "";
}

export function isManualQuarterlyTask(
  task: Pick<QuarterlyTaskItem, "kind" | "source">,
): boolean {
  return !isProtectedQuarterlyTask(task);
}

export function createManualQuarterlyTask(
  label: string,
  dueDate = "",
): QuarterlyTaskItem {
  return {
    id: createQuarterlyTaskId(),
    kind: "",
    label: label.trim(),
    dueDate: dueDate.trim(),
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

export function updateQuarterlyTaskFields(
  tasks: QuarterlyTaskItem[],
  taskId: string,
  patch: { label?: string; dueDate?: string },
): QuarterlyTaskItem[] {
  return tasks.map((task) => {
    if (task.id !== taskId) {
      return task;
    }
    // Default tasks: due date editable; label locked.
    if (isProtectedQuarterlyTask(task)) {
      if (patch.dueDate === undefined) {
        return task;
      }
      return {
        ...task,
        dueDate: patch.dueDate.trim(),
      };
    }
    return {
      ...task,
      label:
        patch.label !== undefined ? patch.label.trim() || task.label : task.label,
      dueDate:
        patch.dueDate !== undefined ? patch.dueDate.trim() : task.dueDate,
    };
  });
}

/** @deprecated Prefer updateQuarterlyTaskFields */
export function updateQuarterlyTaskLabel(
  tasks: QuarterlyTaskItem[],
  taskId: string,
  label: string,
): QuarterlyTaskItem[] {
  return updateQuarterlyTaskFields(tasks, taskId, { label });
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
      dueDate: prior.dueDate || task.dueDate,
      kind: prior.kind,
      source: QUARTERLY_GST_HST_TASK_SOURCE,
      completed: task.completed,
      completedAt: task.completedAt,
      completedByName: task.completedByName,
      completedByEmail: task.completedByEmail,
    };
  });
  return missing.length > 0 ? [...missing, ...merged] : merged;
}

/**
 * When proposal has GST/HST Tax Compliance Service(s) heading, ensure the
 * fixed quarterly checklist. Due dates start empty for users to set.
 */
export function ensureGstHstTaxComplianceQuarterlyTasks(
  tasks: QuarterlyTaskItem[],
  proposalSlug = "",
): QuarterlyTaskItem[] {
  const existingKinds = new Set(
    tasks.filter(isProtectedQuarterlyTask).map((task) => task.kind),
  );
  const additions: QuarterlyTaskItem[] = [];
  const slug = proposalSlug.trim();
  for (const def of GST_HST_TAX_COMPLIANCE_QUARTERLY_TASK_DEFINITIONS) {
    if (existingKinds.has(def.kind)) {
      continue;
    }
    existingKinds.add(def.kind);
    additions.push({
      id: createQuarterlyTaskId(),
      kind: def.kind,
      label: def.label,
      dueDate: "",
      completed: false,
      completedAt: "",
      completedByName: "",
      completedByEmail: "",
      createdAt: new Date().toISOString(),
      source: QUARTERLY_GST_HST_TASK_SOURCE,
      proposalSlug: slug,
    });
  }
  return additions.length > 0 ? [...tasks, ...additions] : tasks;
}
