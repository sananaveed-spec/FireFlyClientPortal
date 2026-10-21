import {
  formatBusinessDayLabel,
  getScheduleBusinessDate,
} from "@/lib/business-days";
import { CLIENT_MANAGEMENT_BD_SLOTS } from "@/lib/client-management-bd";
import {
  classifyHolidayName,
  computeHolidayDate,
  isFloatingHolidayName,
  isLunarHolidayName,
} from "@/lib/holiday-rules";
import {
  normalizeHolidayDate,
  type HolidayCountryId,
  type HolidayEntry,
  type HolidaysStore,
} from "@/lib/holidays";

export type RepeatedTaskStatus =
  | "idle"
  | "due"
  | "running"
  | "done"
  | "failed";

export type RepeatedTaskSchedule = "yearly-on-jan-1" | "monthly-on-1st";

export type RepeatedTaskChecklistItem = {
  id: string;
  label: string;
  /** System step completed (copy/skip). */
  done: boolean;
  performedAt: string | null;
  detail: string;
  /**
   * When true, item used to support Mark checked (legacy).
   * Lunar items use `manual` instead — no check-off.
   */
  requiresHumanCheck: boolean;
  /** Informational only: must be set in Holidays by a person. */
  manual: boolean;
  checkedBy: string | null;
  checkedAt: string | null;
};

/** Per-BD status for the monthly Client Management refresh. */
export type ClientManagementBdRefresh = {
  id: string;
  label: string;
  businessDateLabel: string;
  deadlineRule: string;
  defaultTemplate: string;
  refreshedAt: string | null;
};

export type RepeatedTask = {
  id: string;
  name: string;
  description: string;
  schedule: RepeatedTaskSchedule;
  status: RepeatedTaskStatus;
  lastPerformedAt: string | null;
  nextPerformAt: string | null;
  /** Calendar year the holiday rollover last targeted (e.g. 2027). */
  lastTargetYear: number | null;
  /** YYYY-MM for monthly Client Management refresh. */
  lastTargetMonth: string | null;
  /** Populated for monthly Client Management task. */
  bdSlots: ClientManagementBdRefresh[];
  /**
   * Demo override YYYY-MM for Client Management (null = use real calendar).
   * Only used on the monthly Client Management task.
   */
  demoMonth: string | null;
  checklist: RepeatedTaskChecklistItem[];
  lastError: string | null;
};

export type RepeatedTasksStore = {
  tasks: RepeatedTask[];
};

export const ROLL_HOLIDAYS_TASK_ID = "roll-holidays-new-year";
export const ROLL_CLIENT_MANAGEMENT_MONTH_ID = "roll-client-management-month";

/** Browser event when demo month changes — Client Management listens. */
export const CLIENT_MANAGEMENT_DEMO_MONTH_EVENT =
  "firefly-client-management-demo-month";

/** Mid-month Date for a YYYY-MM key (stable for BD ordinals). */
export function parseMonthKeyToDate(month: string): Date | null {
  if (!/^\d{4}-\d{2}$/.test(month.trim())) {
    return null;
  }
  const [year, mon] = month.trim().split("-").map(Number);
  if (!year || !mon || mon < 1 || mon > 12) {
    return null;
  }
  return new Date(year, mon - 1, 15, 12, 0, 0, 0);
}

/** Clock used for Client Management BD dates (demo month or real now). */
export function getClientManagementNow(
  demoMonth: string | null | undefined,
  realNow: Date = new Date(),
): Date {
  if (!demoMonth) {
    return realNow;
  }
  return parseMonthKeyToDate(demoMonth) ?? realNow;
}

function makeChecklistId(): string {
  return `chk-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function makeHolidayId(prefix: string): string {
  return `holiday-${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function jan1Iso(year: number): string {
  return new Date(year, 0, 1, 0, 0, 0, 0).toISOString();
}

function monthKey(now: Date): string {
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function firstOfNextMonthIso(now: Date): string {
  return new Date(now.getFullYear(), now.getMonth() + 1, 1, 0, 0, 0, 0).toISOString();
}

function firstOfMonthIso(now: Date): string {
  return new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0).toISOString();
}

function stampItem(
  label: string,
  detail: string,
  options: {
    done?: boolean;
    requiresHumanCheck?: boolean;
    manual?: boolean;
    at?: Date;
  } = {},
): RepeatedTaskChecklistItem {
  const at = options.at ?? new Date();
  const manual = Boolean(options.manual);
  const requiresHumanCheck = Boolean(options.requiresHumanCheck) && !manual;
  const done = manual ? false : (options.done ?? !requiresHumanCheck);
  return {
    id: makeChecklistId(),
    label,
    done,
    performedAt: at.toISOString(),
    detail,
    requiresHumanCheck,
    manual,
    checkedBy: null,
    checkedAt: null,
  };
}

export function createDefaultRepeatedTasks(
  now: Date = new Date(),
): RepeatedTasksStore {
  const year = now.getFullYear();
  return {
    tasks: [
      {
        id: ROLL_HOLIDAYS_TASK_ID,
        name: "Roll Holidays to Current Year",
        description:
          "Copies prior-year holidays into the current year (automatic on Lookup open). Floating holidays are auto-set; lunar holidays must be set manually in Holidays.",
        schedule: "yearly-on-jan-1",
        status: "idle",
        lastPerformedAt: null,
        nextPerformAt: jan1Iso(year),
        lastTargetYear: null,
        lastTargetMonth: null,
        bdSlots: [],
        demoMonth: null,
        checklist: [],
        lastError: null,
      },
      {
        id: ROLL_CLIENT_MANAGEMENT_MONTH_ID,
        name: "Roll Client Management to Current Month",
        description:
          "On the 1st of each month (automatic on Lookup open), refreshes Standard Client Management business days, default templates, and due-date rules for the current month. Use Demo month to preview any month.",
        schedule: "monthly-on-1st",
        status: "idle",
        lastPerformedAt: null,
        nextPerformAt: firstOfMonthIso(now),
        lastTargetYear: null,
        lastTargetMonth: null,
        bdSlots: [],
        demoMonth: null,
        checklist: [],
        lastError: null,
      },
    ],
  };
}

function normalizeChecklistItem(value: unknown): RepeatedTaskChecklistItem | null {
  if (typeof value !== "object" || value === null) {
    return null;
  }
  const raw = value as Partial<RepeatedTaskChecklistItem>;
  const label = String(raw.label ?? "").trim();
  if (!label) {
    return null;
  }
  const checkedBy =
    typeof raw.checkedBy === "string" && raw.checkedBy.trim()
      ? raw.checkedBy.trim()
      : null;
  const checkedAt =
    typeof raw.checkedAt === "string" && raw.checkedAt.trim()
      ? raw.checkedAt.trim()
      : null;
  const manual =
    Boolean(raw.manual) ||
    /^lunar holidays$/i.test(label) ||
    /verify lunar/i.test(label);
  const requiresHumanCheck =
    !manual && Boolean(raw.requiresHumanCheck);
  const done = manual ? false : Boolean(raw.done);
  let detail = String(raw.detail ?? "").trim();
  detail = detail
    .replace(/\s*Then mark Checked\.?/gi, "")
    .replace(/\s*then mark checked\.?/gi, "")
    .trim();
  if (
    manual &&
    (/^Update in Holidays if needed:/i.test(detail) ||
      /^Lunar holidays need to be set manually/i.test(detail))
  ) {
    detail = "";
  }
  return {
    id:
      typeof raw.id === "string" && raw.id.trim()
        ? raw.id.trim()
        : makeChecklistId(),
    label,
    done,
    performedAt:
      typeof raw.performedAt === "string" && raw.performedAt.trim()
        ? raw.performedAt.trim()
        : null,
    detail,
    requiresHumanCheck,
    manual,
    checkedBy,
    checkedAt,
  };
}

function normalizeBdSlot(value: unknown): ClientManagementBdRefresh | null {
  if (typeof value !== "object" || value === null) {
    return null;
  }
  const raw = value as Partial<ClientManagementBdRefresh>;
  const id = String(raw.id ?? "").trim();
  const label = String(raw.label ?? "").trim();
  if (!id || !label) {
    return null;
  }
  return {
    id,
    label,
    businessDateLabel: String(raw.businessDateLabel ?? "—").trim() || "—",
    deadlineRule: String(raw.deadlineRule ?? "—").trim() || "—",
    defaultTemplate: String(raw.defaultTemplate ?? "—").trim() || "—",
    refreshedAt:
      typeof raw.refreshedAt === "string" && raw.refreshedAt.trim()
        ? raw.refreshedAt.trim()
        : null,
  };
}

function normalizeTask(value: unknown, fallback: RepeatedTask): RepeatedTask {
  if (typeof value !== "object" || value === null) {
    return fallback;
  }
  const raw = value as Partial<RepeatedTask>;
  const status = raw.status;
  const validStatus: RepeatedTaskStatus[] = [
    "idle",
    "due",
    "running",
    "done",
    "failed",
  ];
  return {
    id: fallback.id,
    // Keep seeded task titles/descriptions in sync with code.
    name: fallback.name,
    description: fallback.description,
    schedule: fallback.schedule,
    status: validStatus.includes(status as RepeatedTaskStatus)
      ? (status as RepeatedTaskStatus)
      : fallback.status,
    lastPerformedAt:
      typeof raw.lastPerformedAt === "string" ? raw.lastPerformedAt : null,
    nextPerformAt:
      typeof raw.nextPerformAt === "string" ? raw.nextPerformAt : null,
    lastTargetYear:
      typeof raw.lastTargetYear === "number" &&
      Number.isFinite(raw.lastTargetYear)
        ? raw.lastTargetYear
        : null,
    lastTargetMonth:
      typeof raw.lastTargetMonth === "string" &&
      /^\d{4}-\d{2}$/.test(raw.lastTargetMonth.trim())
        ? raw.lastTargetMonth.trim()
        : null,
    bdSlots: Array.isArray(raw.bdSlots)
      ? raw.bdSlots
          .map(normalizeBdSlot)
          .filter((item): item is ClientManagementBdRefresh => item !== null)
      : [],
    demoMonth:
      fallback.id === ROLL_CLIENT_MANAGEMENT_MONTH_ID &&
      typeof raw.demoMonth === "string" &&
      /^\d{4}-\d{2}$/.test(raw.demoMonth.trim())
        ? raw.demoMonth.trim()
        : null,
    checklist: Array.isArray(raw.checklist)
      ? raw.checklist
          .map(normalizeChecklistItem)
          .filter((item): item is RepeatedTaskChecklistItem => item !== null)
      : [],
    lastError:
      typeof raw.lastError === "string" && raw.lastError.trim()
        ? raw.lastError.trim()
        : null,
  };
}

export function normalizeRepeatedTasks(value: unknown): RepeatedTasksStore {
  const defaults = createDefaultRepeatedTasks();
  const incoming =
    typeof value === "object" && value !== null
      ? (value as { tasks?: unknown })
      : {};
  const incomingTasks = Array.isArray(incoming.tasks) ? incoming.tasks : [];
  const byId = new Map<string, unknown>();
  for (const task of incomingTasks) {
    if (typeof task === "object" && task !== null) {
      const id = String((task as { id?: unknown }).id ?? "").trim();
      if (id) {
        byId.set(id, task);
      }
    }
  }

  return {
    tasks: defaults.tasks.map((fallback) =>
      normalizeTask(byId.get(fallback.id), fallback),
    ),
  };
}

function holidayYear(date: string): number | null {
  const match = normalizeHolidayDate(date).match(/^(\d{4})-/);
  if (!match) {
    return null;
  }
  return Number(match[1]);
}

function bumpHolidayYear(date: string, targetYear: number): string | null {
  const normalized = normalizeHolidayDate(date);
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return null;
  }
  return `${targetYear}-${match[2]}-${match[3]}`;
}

function countryHasYear(
  holidays: HolidayEntry[],
  year: number,
): boolean {
  return holidays.some((h) => holidayYear(h.date) === year);
}

function sourceYearForCountry(
  holidays: HolidayEntry[],
  targetYear: number,
): number | null {
  const years = new Set<number>();
  for (const h of holidays) {
    const y = holidayYear(h.date);
    if (y !== null && y < targetYear) {
      years.add(y);
    }
  }
  if (years.size === 0) {
    return null;
  }
  return Math.max(...years);
}

function resolveDateForYear(
  name: string,
  countryId: HolidayCountryId,
  targetYear: number,
  sourceDate: string,
): string | null {
  const computed = computeHolidayDate(name, countryId, targetYear);
  if (computed) {
    return computed;
  }
  return bumpHolidayYear(sourceDate, targetYear);
}

function copyHolidaysToYear(
  entries: HolidayEntry[],
  sourceYear: number,
  targetYear: number,
  countryId: HolidayCountryId,
): HolidayEntry[] {
  return entries
    .filter((h) => holidayYear(h.date) === sourceYear)
    .map((h) => {
      const nextDate = resolveDateForYear(
        h.name,
        countryId,
        targetYear,
        h.date,
      );
      if (!nextDate) {
        return null;
      }
      const entry: HolidayEntry = {
        id: makeHolidayId(countryId),
        name: h.name,
        date: nextDate,
        // Year-shift placeholder until someone edits (esp. lunar dates).
        origin: "auto",
      };
      if (h.state !== undefined) {
        entry.state = h.state;
      }
      return entry;
    })
    .filter((item): item is HolidayEntry => item !== null);
}

/** Recompute floating / rule-based dates for an existing year in place. */
function applyComputedDatesForYear(
  entries: HolidayEntry[],
  countryId: HolidayCountryId,
  year: number,
): { floatingUpdated: string[]; lunarNames: string[] } {
  const floatingUpdated: string[] = [];
  const lunarNames: string[] = [];
  for (const entry of entries) {
    if (holidayYear(entry.date) !== year) {
      continue;
    }
    const kind = classifyHolidayName(entry.name, countryId);
    if (kind === "lunar") {
      lunarNames.push(entry.name);
      // Untagged / system-copied lunar → auto until a person saves an edit.
      if (entry.origin !== "manual") {
        entry.origin = "auto";
      }
      continue;
    }
    const computed = computeHolidayDate(entry.name, countryId, year);
    if (!computed) {
      continue;
    }
    if (entry.date !== computed) {
      entry.date = computed;
    }
    if (isFloatingHolidayName(entry.name, countryId)) {
      floatingUpdated.push(`${entry.name} → ${computed}`);
    }
  }
  return { floatingUpdated, lunarNames };
}

export type EnsureRepeatedTasksResult = {
  tasks: RepeatedTasksStore;
  holidays: HolidaysStore;
  ran: boolean;
  message: string;
};

function markHolidayTaskDone(
  task: RepeatedTask,
  targetYear: number,
  checklist: RepeatedTaskChecklistItem[],
  at: Date,
): RepeatedTask {
  return {
    ...task,
    status: "done",
    lastPerformedAt: at.toISOString(),
    nextPerformAt: jan1Iso(targetYear + 1),
    lastTargetYear: targetYear,
    checklist,
    lastError: null,
  };
}

function markMonthTaskDone(
  task: RepeatedTask,
  targetMonth: string,
  bdSlots: ClientManagementBdRefresh[],
  checklist: RepeatedTaskChecklistItem[],
  at: Date,
  now: Date,
): RepeatedTask {
  return {
    ...task,
    status: "done",
    lastPerformedAt: at.toISOString(),
    nextPerformAt: firstOfNextMonthIso(now),
    lastTargetMonth: targetMonth,
    bdSlots,
    checklist,
    lastError: null,
  };
}

function buildClientManagementBdSlots(
  holidays: HolidaysStore,
  now: Date,
  at: Date,
): ClientManagementBdRefresh[] {
  return CLIENT_MANAGEMENT_BD_SLOTS.map((slot) => {
    const date = getScheduleBusinessDate(
      slot.id,
      holidays,
      undefined,
      "",
      now,
    );
    return {
      id: slot.id,
      label: slot.label,
      businessDateLabel: date ? formatBusinessDayLabel(date) : "—",
      deadlineRule: slot.deadlineRule,
      defaultTemplate: slot.defaultTemplate,
      refreshedAt: at.toISOString(),
    };
  });
}

/**
 * Idempotent: if the current calendar year is not yet stamped on the task,
 * copy missing country holidays from the prior year and stamp the checklist.
 * Safe to call on every Lookup open.
 */
export function ensureRepeatedTasks(
  tasksStore: RepeatedTasksStore,
  holidaysStore: HolidaysStore,
  now: Date = new Date(),
): EnsureRepeatedTasksResult {
  const normalizedTasks = normalizeRepeatedTasks(tasksStore);
  const targetYear = now.getFullYear();
  const at = now;
  let holidays = holidaysStore;
  let ran = false;
  let message = "No due repeated tasks.";

  const tasks = normalizedTasks.tasks.map((task) => {
    if (task.id === ROLL_CLIENT_MANAGEMENT_MONTH_ID) {
      const clock = getClientManagementNow(task.demoMonth, now);
      const currentMonth = monthKey(clock);
      if (
        task.lastTargetMonth === currentMonth &&
        task.status === "done" &&
        task.bdSlots.length > 0
      ) {
        return {
          ...task,
          nextPerformAt: firstOfNextMonthIso(clock),
          lastError: null,
        };
      }

      ran = true;
      const bdSlots = buildClientManagementBdSlots(holidays, clock, at);
      const checklist: RepeatedTaskChecklistItem[] = [
        stampItem(
          `Month ${currentMonth} refreshed`,
          task.demoMonth
            ? `Demo month ${task.demoMonth}: business day dates, default templates, and due-date rules.`
            : "Business day dates, default templates, and due-date rules updated for Standard Client Management.",
          { done: true, at },
        ),
        ...bdSlots.map((slot) =>
          stampItem(
            `${slot.label}: dates & due rules`,
            `Date: ${slot.businessDateLabel}. Deadline: ${slot.deadlineRule}. Template: ${slot.defaultTemplate}.`,
            { done: true, at },
          ),
        ),
      ];
      message = task.demoMonth
        ? `Client Management demo month set to ${currentMonth}.`
        : `Client Management refreshed for ${currentMonth}.`;
      return {
        ...markMonthTaskDone(
          task,
          currentMonth,
          bdSlots,
          checklist,
          at,
          clock,
        ),
        demoMonth: task.demoMonth,
      };
    }

    if (task.id !== ROLL_HOLIDAYS_TASK_ID) {
      return task;
    }

    const hasSplitChecklist = task.checklist.some((item) =>
      /^floating holidays/i.test(item.label),
    );
    const hasManualLunar = task.checklist.some(
      (item) => /^lunar holidays$/i.test(item.label) && item.manual,
    );
    if (
      task.lastTargetYear === targetYear &&
      task.status === "done" &&
      hasSplitChecklist &&
      hasManualLunar
    ) {
      // Stamp untagged lunar rows as auto (amber) without rebuilding the whole checklist.
      const working: HolidaysStore = {
        countries: holidays.countries.map((country) => ({
          ...country,
          holidays: country.holidays.map((entry) => ({ ...entry })),
        })),
      };
      let stamped = false;
      for (const country of working.countries) {
        for (const entry of country.holidays) {
          if (
            holidayYear(entry.date) === targetYear &&
            isLunarHolidayName(entry.name) &&
            entry.origin !== "manual" &&
            entry.origin !== "auto"
          ) {
            entry.origin = "auto";
            stamped = true;
          }
        }
      }
      if (stamped) {
        holidays = working;
        ran = true;
        message = `Marked lunar holidays for ${targetYear} as auto until edited.`;
      }
      return {
        ...task,
        checklist: task.checklist.map((item) => ({ ...item })),
        nextPerformAt: jan1Iso(targetYear + 1),
        lastError: null,
      };
    }

    ran = true;
    const checklist: RepeatedTaskChecklistItem[] = [];
    const working: HolidaysStore = {
      countries: holidays.countries.map((country) => ({
        ...country,
        holidays: [...country.holidays],
      })),
    };

    let copiedAny = false;
    const allFloatingNotes: string[] = [];

    for (const country of working.countries) {
      if (!countryHasYear(country.holidays, targetYear)) {
        const sourceYear = sourceYearForCountry(
          country.holidays,
          targetYear,
        );
        if (sourceYear === null) {
          checklist.push(
            stampItem(
              `${country.name}: no prior year to copy`,
              `No holidays before ${targetYear} found. Add dates manually in Holidays.`,
              { done: true, at },
            ),
          );
        } else {
          const copied = copyHolidaysToYear(
            country.holidays,
            sourceYear,
            targetYear,
            country.id,
          );
          country.holidays = [...country.holidays, ...copied];
          copiedAny = copiedAny || copied.length > 0;
          checklist.push(
            stampItem(
              `${country.name}: copied ${sourceYear} → ${targetYear}`,
              `Added ${copied.length} holiday${copied.length === 1 ? "" : "s"} (fixed/lunar year-shift; floating dates computed).`,
              { done: true, at },
            ),
          );
        }
      } else {
        checklist.push(
          stampItem(
            `${country.name}: already has ${targetYear}`,
            "Dates already exist for this year — recomputed floating rules.",
            { done: true, at },
          ),
        );
      }

      const applied = applyComputedDatesForYear(
        country.holidays,
        country.id,
        targetYear,
      );
      for (const note of applied.floatingUpdated) {
        allFloatingNotes.push(`${country.name}: ${note}`);
      }
      // Also list floating that were already correct (still auto-set).
      for (const entry of country.holidays) {
        if (holidayYear(entry.date) !== targetYear) {
          continue;
        }
        if (isFloatingHolidayName(entry.name, country.id)) {
          const key = `${country.name}: ${entry.name} → ${entry.date}`;
          if (!allFloatingNotes.includes(key)) {
            allFloatingNotes.push(key);
          }
        }
      }
    }

    checklist.push(
      stampItem(
        "Floating holidays: auto-set by system",
        allFloatingNotes.length > 0
          ? allFloatingNotes.join("; ")
          : "No floating holidays found for this year.",
        { done: true, at },
      ),
    );

    checklist.push(
      stampItem("Lunar holidays", "", { done: false, manual: true, at }),
    );

    holidays = working;
    message = copiedAny
      ? `Rolled holidays into ${targetYear}.`
      : `Holidays for ${targetYear} updated (floating auto-set).`;
    return markHolidayTaskDone(task, targetYear, checklist, at);
  });

  return {
    tasks: { tasks },
    holidays,
    ran,
    message,
  };
}

/** Force-run a seeded task (admin recovery). */
export function runRepeatedTask(
  taskId: string,
  tasksStore: RepeatedTasksStore,
  holidaysStore: HolidaysStore,
  now: Date = new Date(),
): EnsureRepeatedTasksResult {
  if (
    taskId !== ROLL_HOLIDAYS_TASK_ID &&
    taskId !== ROLL_CLIENT_MANAGEMENT_MONTH_ID
  ) {
    throw new Error("Unknown repeated task.");
  }
  const normalized = normalizeRepeatedTasks(tasksStore);
  const cleared: RepeatedTasksStore = {
    tasks: normalized.tasks.map((task) => {
      if (task.id === ROLL_HOLIDAYS_TASK_ID && taskId === ROLL_HOLIDAYS_TASK_ID) {
        return {
          ...task,
          status: "due" as const,
          lastTargetYear: null,
          lastError: null,
        };
      }
      if (
        task.id === ROLL_CLIENT_MANAGEMENT_MONTH_ID &&
        taskId === ROLL_CLIENT_MANAGEMENT_MONTH_ID
      ) {
        return {
          ...task,
          status: "due" as const,
          lastTargetMonth: null,
          bdSlots: [],
          lastError: null,
        };
      }
      return task;
    }),
  };
  return ensureRepeatedTasks(cleared, holidaysStore, now);
}

/**
 * Set or clear demo month for Client Management, then rebuild BD slots for that month.
 */
export function setClientManagementDemoMonth(
  tasksStore: RepeatedTasksStore,
  holidaysStore: HolidaysStore,
  demoMonth: string | null,
  now: Date = new Date(),
): EnsureRepeatedTasksResult {
  const normalized = normalizeRepeatedTasks(tasksStore);
  const nextDemo =
    demoMonth && /^\d{4}-\d{2}$/.test(demoMonth.trim())
      ? demoMonth.trim()
      : null;
  const prepared: RepeatedTasksStore = {
    tasks: normalized.tasks.map((task) =>
      task.id === ROLL_CLIENT_MANAGEMENT_MONTH_ID
        ? {
            ...task,
            demoMonth: nextDemo,
            lastTargetMonth: null,
            bdSlots: [],
            status: "due",
            lastError: null,
          }
        : task,
    ),
  };
  return ensureRepeatedTasks(prepared, holidaysStore, now);
}

/** Mark a human-review checklist item as checked or remaining. */
export function setChecklistItemReview(
  tasksStore: RepeatedTasksStore,
  taskId: string,
  itemId: string,
  checked: boolean,
  checkedBy: string,
  now: Date = new Date(),
): RepeatedTasksStore {
  const normalized = normalizeRepeatedTasks(tasksStore);
  const who = checkedBy.trim() || "Unknown";
  return {
    tasks: normalized.tasks.map((task) => {
      if (task.id !== taskId) {
        return task;
      }
      return {
        ...task,
        checklist: task.checklist.map((item) => {
          if (item.id !== itemId || !item.requiresHumanCheck) {
            return item;
          }
          if (checked) {
            return {
              ...item,
              done: true,
              checkedBy: who,
              checkedAt: now.toISOString(),
            };
          }
          return {
            ...item,
            done: false,
            checkedBy: null,
            checkedAt: null,
          };
        }),
      };
    }),
  };
}
