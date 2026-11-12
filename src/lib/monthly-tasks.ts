/**
 * Client Info — Monthly Tasks checklist (Ignition monthly / recurring services).
 */

export const MONTHLY_SERVICE_TASK_SOURCE = "ignition_monthly" as const;

/** Fixed Monthly Tasks when proposal has "Bookkeeping Services" heading. */
export const BOOKKEEPING_MONTHLY_TASK_DEFINITIONS = [
  {
    kind: "bookkeeping_monthly:bank-feed",
    label:
      "Regular bank and credit card feed monitoring, review, and clearance.",
  },
  {
    kind: "bookkeeping_monthly:reconciliations",
    label: "Monthly bank and credit card account reconciliations.",
  },
  {
    kind: "bookkeeping_monthly:sharepoint-drive",
    label:
      "Monthly SharePoint / Google Drive maintenance and document organization.",
  },
  {
    kind: "bookkeeping_monthly:payroll-journals",
    label:
      "Biweekly / monthly payroll journal entries recording and posting.",
  },
  {
    kind: "bookkeeping_monthly:customer-invoices",
    label: "Preparation and issuance of regular customer invoices.",
  },
  {
    kind: "bookkeeping_monthly:vendor-bills",
    label: "Regular vendor bills uploading, recording, and categorization.",
  },
] as const;

/** Fixed Monthly Tasks for Financial Reporting Service(s) heading. */
export const FINANCIAL_REPORTING_MONTHLY_TASK_DEFINITIONS = [
  {
    kind: "financial_reporting_monthly:ar-reports",
    label: "Weekly / monthly Accounts Receivable (AR) reports.",
  },
  {
    kind: "financial_reporting_monthly:ap-reports",
    label: "Weekly / monthly Accounts Payable (AP) reports.",
  },
  {
    kind: "financial_reporting_monthly:pl-statements",
    label: "Monthly / quarterly Profit & Loss (P&L) statements.",
  },
  {
    kind: "financial_reporting_monthly:balance-sheet",
    label: "Monthly / quarterly Balance Sheet statements.",
  },
  {
    kind: "financial_reporting_monthly:cash-flow",
    label: "Monthly / quarterly Cash Flow statements.",
  },
  {
    kind: "financial_reporting_monthly:budget-variance",
    label: "Monthly / quarterly budget and variance analysis.",
  },
] as const;

export type MonthlyTaskItem = {
  id: string;
  /** Stable key for default monthly services; empty for user-added. */
  kind: string;
  label: string;
  completed: boolean;
  completedAt: string;
  completedByName: string;
  completedByEmail: string;
  createdAt: string;
  source: typeof MONTHLY_SERVICE_TASK_SOURCE | "manual";
  proposalSlug: string;
};

export function createMonthlyTaskId() {
  return `mot-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function monthlyServiceTaskKind(serviceName: string): string {
  const slug = serviceName
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
  return `monthly_service:${slug || "item"}`;
}

/** Ignition configured_as / billing_mode looks like monthly recurring. */
export function isMonthlyBillingMode(billingMode: string): boolean {
  const mode = billingMode.trim().toLowerCase().replace(/[_-]+/g, " ");
  if (!mode) {
    return false;
  }
  if (/once\s*off|one[\s-]*time|annual|yearly|quarter|week/.test(mode)) {
    return false;
  }
  return /month/.test(mode) || mode === "recurring";
}

export function normalizeMonthlyTaskItem(
  value: unknown,
): MonthlyTaskItem | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<MonthlyTaskItem>;
  const id = String(record.id ?? "").trim();
  const label = String(record.label ?? "").trim();
  if (!id || !label) {
    return null;
  }
  const source =
    record.source === MONTHLY_SERVICE_TASK_SOURCE
      ? MONTHLY_SERVICE_TASK_SOURCE
      : "manual";
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

export function normalizeMonthlyTasks(value: unknown): MonthlyTaskItem[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map(normalizeMonthlyTaskItem)
    .filter((item): item is MonthlyTaskItem => Boolean(item));
}

export function monthlyTasksStatusLabel(tasks: MonthlyTaskItem[]): string {
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

export function isProtectedMonthlyTask(
  task: Pick<MonthlyTaskItem, "kind" | "source">,
): boolean {
  const kind = String(task.kind ?? "");
  return (
    task.source === MONTHLY_SERVICE_TASK_SOURCE ||
    kind.startsWith("monthly_service:") ||
    kind.startsWith("bookkeeping_monthly:") ||
    kind.startsWith("financial_reporting_monthly:")
  );
}

export function isManualMonthlyTask(
  task: Pick<MonthlyTaskItem, "kind" | "source">,
): boolean {
  return !isProtectedMonthlyTask(task);
}

export function createManualMonthlyTask(label: string): MonthlyTaskItem {
  return {
    id: createMonthlyTaskId(),
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

export function createMonthlyServiceTask(input: {
  serviceName: string;
  proposalSlug?: string;
}): MonthlyTaskItem {
  const label = input.serviceName.trim();
  return {
    id: createMonthlyTaskId(),
    kind: monthlyServiceTaskKind(label),
    label,
    completed: false,
    completedAt: "",
    completedByName: "",
    completedByEmail: "",
    createdAt: new Date().toISOString(),
    source: MONTHLY_SERVICE_TASK_SOURCE,
    proposalSlug: (input.proposalSlug || "").trim(),
  };
}

export function toggleMonthlyTaskCompletion(
  tasks: MonthlyTaskItem[],
  taskId: string,
  completed: boolean,
  actor: { name: string; email: string },
): MonthlyTaskItem[] {
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

export function updateMonthlyTaskLabel(
  tasks: MonthlyTaskItem[],
  taskId: string,
  label: string,
): MonthlyTaskItem[] {
  const nextLabel = label.trim();
  return tasks.map((task) => {
    if (task.id !== taskId || isProtectedMonthlyTask(task) || !nextLabel) {
      return task;
    }
    return { ...task, label: nextLabel };
  });
}

export function removeMonthlyTask(
  tasks: MonthlyTaskItem[],
  taskId: string,
): MonthlyTaskItem[] {
  return tasks.filter((task) => {
    if (task.id !== taskId) {
      return true;
    }
    return isProtectedMonthlyTask(task);
  });
}

export function mergeProtectedMonthlyTasks(
  previous: MonthlyTaskItem[],
  next: MonthlyTaskItem[],
): MonthlyTaskItem[] {
  const protectedPrevious = previous.filter(isProtectedMonthlyTask);
  const nextIds = new Set(next.map((task) => task.id));
  const nextKinds = new Set(
    next.filter(isProtectedMonthlyTask).map((task) => task.kind),
  );
  const missing = protectedPrevious.filter(
    (task) => !nextIds.has(task.id) && !nextKinds.has(task.kind),
  );
  const merged = next.map((task) => {
    if (!isProtectedMonthlyTask(task)) {
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
      source: MONTHLY_SERVICE_TASK_SOURCE,
      completed: task.completed,
      completedAt: task.completedAt,
      completedByName: task.completedByName,
      completedByEmail: task.completedByEmail,
    };
  });
  return missing.length > 0 ? [...missing, ...merged] : merged;
}

export function ensureMonthlyTasksFromServices(
  tasks: MonthlyTaskItem[],
  monthlyServices: ReadonlyArray<{ name: string }>,
  proposalSlug = "",
): MonthlyTaskItem[] {
  if (monthlyServices.length === 0) {
    return tasks;
  }
  const existingKinds = new Set(
    tasks.filter(isProtectedMonthlyTask).map((task) => task.kind),
  );
  const additions: MonthlyTaskItem[] = [];
  for (const service of monthlyServices) {
    const name = service.name.trim();
    if (!name) {
      continue;
    }
    const kind = monthlyServiceTaskKind(name);
    if (existingKinds.has(kind)) {
      continue;
    }
    existingKinds.add(kind);
    additions.push(
      createMonthlyServiceTask({
        serviceName: name,
        proposalSlug,
      }),
    );
  }
  return additions.length > 0 ? [...tasks, ...additions] : tasks;
}

function ensureFixedMonthlyTasksFromDefinitions(
  tasks: MonthlyTaskItem[],
  definitions: ReadonlyArray<{ kind: string; label: string }>,
  proposalSlug = "",
): MonthlyTaskItem[] {
  const existingKinds = new Set(
    tasks.filter(isProtectedMonthlyTask).map((task) => task.kind),
  );
  const additions: MonthlyTaskItem[] = [];
  const slug = proposalSlug.trim();
  for (const def of definitions) {
    if (existingKinds.has(def.kind)) {
      continue;
    }
    existingKinds.add(def.kind);
    additions.push({
      id: createMonthlyTaskId(),
      kind: def.kind,
      label: def.label,
      completed: false,
      completedAt: "",
      completedByName: "",
      completedByEmail: "",
      createdAt: new Date().toISOString(),
      source: MONTHLY_SERVICE_TASK_SOURCE,
      proposalSlug: slug,
    });
  }
  return additions.length > 0 ? [...tasks, ...additions] : tasks;
}

/**
 * When proposal has Bookkeeping Services heading, ensure the fixed monthly
 * checklist (no duplicates by kind).
 */
export function ensureBookkeepingMonthlyTasks(
  tasks: MonthlyTaskItem[],
  proposalSlug = "",
): MonthlyTaskItem[] {
  return ensureFixedMonthlyTasksFromDefinitions(
    tasks,
    BOOKKEEPING_MONTHLY_TASK_DEFINITIONS,
    proposalSlug,
  );
}

/**
 * When proposal has Financial Reporting Service(s) heading, ensure the fixed
 * monthly reporting checklist (no duplicates by kind).
 */
export function ensureFinancialReportingMonthlyTasks(
  tasks: MonthlyTaskItem[],
  proposalSlug = "",
): MonthlyTaskItem[] {
  return ensureFixedMonthlyTasksFromDefinitions(
    tasks,
    FINANCIAL_REPORTING_MONTHLY_TASK_DEFINITIONS,
    proposalSlug,
  );
}
