/**
 * Client Info — Annual Tasks checklist.
 * Income Tax Compliance Service(s) proposal heading seeds fixed bullets;
 * due dates are set by assigned users (not auto-parsed from proposal text).
 */

export const ANNUAL_TAX_TASK_SOURCE = "ignition_tax" as const;

/** Fixed Annual Tasks for Income Tax Compliance Service(s) heading. */
export const INCOME_TAX_COMPLIANCE_ANNUAL_TASK_DEFINITIONS = [
  {
    kind: "income_tax_compliance:corporate-return",
    label: "Annual corporate income tax return preparation and filing.",
  },
  {
    kind: "income_tax_compliance:personal-return",
    label: "Annual personal income tax return preparation and filing.",
  },
] as const;

export type AnnualTaskItem = {
  id: string;
  /** Stable key for default tax services; empty for user-added. */
  kind: string;
  label: string;
  /** ISO date YYYY-MM-DD when set by an assigned user. */
  dueDate: string;
  completed: boolean;
  completedAt: string;
  completedByName: string;
  completedByEmail: string;
  createdAt: string;
  source: typeof ANNUAL_TAX_TASK_SOURCE | "manual";
  proposalSlug: string;
};

export function createAnnualTaskId() {
  return `ant-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function formatAnnualDueDateLabel(dueDate: string): string {
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

function incomeTaxComplianceKindForLabel(label: string): string {
  const normalized = label.trim().toLowerCase().replace(/\s+/g, " ");
  const match = INCOME_TAX_COMPLIANCE_ANNUAL_TASK_DEFINITIONS.find(
    (def) => def.label.toLowerCase() === normalized,
  );
  return match?.kind ?? "";
}

export function normalizeAnnualTaskItem(value: unknown): AnnualTaskItem | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<AnnualTaskItem>;
  const id = String(record.id ?? "").trim();
  const label = String(record.label ?? "").trim();
  if (!id || !label) {
    return null;
  }
  const kindFromLabel = incomeTaxComplianceKindForLabel(label);
  const kind = String(record.kind ?? "").trim() || kindFromLabel;
  const isIncomeTaxDefault = kind.startsWith("income_tax_compliance:");
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
      isIncomeTaxDefault || record.source === ANNUAL_TAX_TASK_SOURCE
        ? ANNUAL_TAX_TASK_SOURCE
        : "manual",
    proposalSlug: String(record.proposalSlug ?? "").trim(),
  };
}

/** Old auto-seed from service names (e.g. "Accounting & Tax services"). */
export function isLegacyTaxServiceAnnualTask(
  task: Pick<AnnualTaskItem, "kind">,
): boolean {
  return String(task.kind ?? "").startsWith("tax_service:");
}

export function normalizeAnnualTasks(value: unknown): AnnualTaskItem[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map(normalizeAnnualTaskItem)
    .filter((item): item is AnnualTaskItem => Boolean(item))
    .filter((item) => !isLegacyTaxServiceAnnualTask(item));
}

export function annualTasksStatusLabel(tasks: AnnualTaskItem[]): string {
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

export function isProtectedAnnualTask(
  task: Pick<AnnualTaskItem, "kind" | "source">,
): boolean {
  return String(task.kind ?? "").startsWith("income_tax_compliance:");
}

export function isManualAnnualTask(
  task: Pick<AnnualTaskItem, "kind" | "source">,
): boolean {
  return !isProtectedAnnualTask(task);
}

export function createManualAnnualTask(
  label: string,
  dueDate = "",
): AnnualTaskItem {
  return {
    id: createAnnualTaskId(),
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

export function toggleAnnualTaskCompletion(
  tasks: AnnualTaskItem[],
  taskId: string,
  completed: boolean,
  actor: { name: string; email: string },
): AnnualTaskItem[] {
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

export function updateAnnualTaskFields(
  tasks: AnnualTaskItem[],
  taskId: string,
  patch: { label?: string; dueDate?: string },
): AnnualTaskItem[] {
  return tasks.map((task) => {
    if (task.id !== taskId) {
      return task;
    }
    // Default (protected) tasks: due date editable by assigned users; label locked.
    if (isProtectedAnnualTask(task)) {
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

export function removeAnnualTask(
  tasks: AnnualTaskItem[],
  taskId: string,
): AnnualTaskItem[] {
  return tasks.filter((task) => {
    if (task.id !== taskId) {
      return true;
    }
    return isProtectedAnnualTask(task);
  });
}

export function mergeProtectedAnnualTasks(
  previous: AnnualTaskItem[],
  next: AnnualTaskItem[],
): AnnualTaskItem[] {
  const protectedPrevious = previous.filter(isProtectedAnnualTask);
  const nextIds = new Set(next.map((task) => task.id));
  const nextKinds = new Set(
    next.filter(isProtectedAnnualTask).map((task) => task.kind),
  );
  const missing = protectedPrevious.filter(
    (task) => !nextIds.has(task.id) && !nextKinds.has(task.kind),
  );
  const merged = next.map((task) => {
    if (!isProtectedAnnualTask(task)) {
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
      source: ANNUAL_TAX_TASK_SOURCE,
      completed: task.completed,
      completedAt: task.completedAt,
      completedByName: task.completedByName,
      completedByEmail: task.completedByEmail,
    };
  });
  return missing.length > 0 ? [...missing, ...merged] : merged;
}

/**
 * When proposal has Income Tax Compliance Service(s) heading, ensure the fixed
 * annual checklist. Due dates start empty for assigned users to set.
 */
export function ensureIncomeTaxComplianceAnnualTasks(
  tasks: AnnualTaskItem[],
  proposalSlug = "",
): AnnualTaskItem[] {
  const existingKinds = new Set(
    tasks.filter(isProtectedAnnualTask).map((task) => task.kind),
  );
  const additions: AnnualTaskItem[] = [];
  const slug = proposalSlug.trim();
  for (const def of INCOME_TAX_COMPLIANCE_ANNUAL_TASK_DEFINITIONS) {
    if (existingKinds.has(def.kind)) {
      continue;
    }
    existingKinds.add(def.kind);
    additions.push({
      id: createAnnualTaskId(),
      kind: def.kind,
      label: def.label,
      dueDate: "",
      completed: false,
      completedAt: "",
      completedByName: "",
      completedByEmail: "",
      createdAt: new Date().toISOString(),
      source: ANNUAL_TAX_TASK_SOURCE,
      proposalSlug: slug,
    });
  }
  return additions.length > 0 ? [...tasks, ...additions] : tasks;
}
