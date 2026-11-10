/**
 * Client Info — Annual Tasks checklist (Ignition Tax / Tax Compliance services).
 */

export const ANNUAL_TAX_TASK_SOURCE = "ignition_tax" as const;

export type AnnualTaskItem = {
  id: string;
  /** Stable key for default tax services; empty for user-added. */
  kind: string;
  label: string;
  /** ISO date YYYY-MM-DD when known. */
  dueDate: string;
  completed: boolean;
  completedAt: string;
  completedByName: string;
  completedByEmail: string;
  createdAt: string;
  source: typeof ANNUAL_TAX_TASK_SOURCE | "manual";
  proposalSlug: string;
};

const MONTH_INDEX: Record<string, number> = {
  january: 1,
  february: 2,
  march: 3,
  april: 4,
  may: 5,
  june: 6,
  july: 7,
  august: 8,
  september: 9,
  october: 10,
  november: 11,
  december: 12,
};

export function createAnnualTaskId() {
  return `ant-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function annualTaxTaskKind(serviceName: string): string {
  const slug = serviceName
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
  return `tax_service:${slug || "item"}`;
}

export function isTaxServiceGroupName(groupName: string): boolean {
  const name = groupName.trim().toLowerCase();
  if (!name) {
    return false;
  }
  return (
    /tax\s*compliance/.test(name) ||
    /tax\s*services?/.test(name) ||
    name === "tax" ||
    /^tax\b/.test(name)
  );
}

/** Strip tags for due-date parsing. */
export function stripHtmlToText(html: string): string {
  return String(html ?? "")
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<\/p>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Parse due dates like "15th of March 2027", "March 15, 2027", "15 March 2027".
 */
export function parseDueDateFromText(text: string): string {
  const raw = stripHtmlToText(text);
  if (!raw) {
    return "";
  }

  const monthNames = Object.keys(MONTH_INDEX).join("|");

  const patterns = [
    new RegExp(
      `due date is\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+of\\s+(${monthNames})\\s+(\\d{4})`,
      "i",
    ),
    new RegExp(
      `due\\s*(?:date)?\\s*(?:is|:)?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+(${monthNames})\\s+(\\d{4})`,
      "i",
    ),
    new RegExp(
      `(${monthNames})\\s+(\\d{1,2})(?:st|nd|rd|th)?(?:,)?\\s+(\\d{4})`,
      "i",
    ),
    /(\d{4})-(\d{2})-(\d{2})/,
  ];

  for (const pattern of patterns) {
    const match = raw.match(pattern);
    if (!match) {
      continue;
    }
    if (pattern.source.startsWith("(\\d{4})")) {
      return `${match[1]}-${match[2]}-${match[3]}`;
    }
    if (MONTH_INDEX[String(match[1] ?? "").toLowerCase()]) {
      const month = MONTH_INDEX[String(match[1]).toLowerCase()]!;
      const day = Number(match[2]);
      const year = Number(match[3]);
      if (day >= 1 && day <= 31 && year >= 2000) {
        return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      }
      continue;
    }
    const day = Number(match[1]);
    const month = MONTH_INDEX[String(match[2] ?? "").toLowerCase()];
    const year = Number(match[3]);
    if (month && day >= 1 && day <= 31 && year >= 2000) {
      return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  return "";
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
  const source =
    record.source === ANNUAL_TAX_TASK_SOURCE ? ANNUAL_TAX_TASK_SOURCE : "manual";
  return {
    id,
    kind: String(record.kind ?? "").trim(),
    label,
    dueDate: String(record.dueDate ?? "").trim(),
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

export function normalizeAnnualTasks(value: unknown): AnnualTaskItem[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map(normalizeAnnualTaskItem)
    .filter((item): item is AnnualTaskItem => Boolean(item));
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
  return (
    task.source === ANNUAL_TAX_TASK_SOURCE ||
    String(task.kind ?? "").startsWith("tax_service:")
  );
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

export function createTaxServiceAnnualTask(input: {
  serviceName: string;
  descriptionHtml?: string;
  proposalSlug?: string;
}): AnnualTaskItem {
  const label = input.serviceName.trim();
  return {
    id: createAnnualTaskId(),
    kind: annualTaxTaskKind(label),
    label,
    dueDate: parseDueDateFromText(input.descriptionHtml || ""),
    completed: false,
    completedAt: "",
    completedByName: "",
    completedByEmail: "",
    createdAt: new Date().toISOString(),
    source: ANNUAL_TAX_TASK_SOURCE,
    proposalSlug: (input.proposalSlug || "").trim(),
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
    if (task.id !== taskId || isProtectedAnnualTask(task)) {
      return task;
    }
    return {
      ...task,
      label: patch.label !== undefined ? patch.label.trim() || task.label : task.label,
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
 * Append missing default tax-service annual tasks from Ignition services.
 * Does not remove or rewrite existing protected tasks (preserves completion).
 */
export function ensureAnnualTasksFromTaxServices(
  tasks: AnnualTaskItem[],
  taxServices: ReadonlyArray<{ name: string; descriptionHtml?: string }>,
  proposalSlug = "",
): AnnualTaskItem[] {
  if (taxServices.length === 0) {
    return tasks;
  }
  const existingKinds = new Set(
    tasks.filter(isProtectedAnnualTask).map((task) => task.kind),
  );
  const additions: AnnualTaskItem[] = [];
  for (const service of taxServices) {
    const name = service.name.trim();
    if (!name) {
      continue;
    }
    const kind = annualTaxTaskKind(name);
    if (existingKinds.has(kind)) {
      continue;
    }
    existingKinds.add(kind);
    additions.push(
      createTaxServiceAnnualTask({
        serviceName: name,
        descriptionHtml: service.descriptionHtml,
        proposalSlug,
      }),
    );
  }
  return additions.length > 0 ? [...tasks, ...additions] : tasks;
}
