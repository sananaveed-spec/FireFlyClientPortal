import type { ClientCountry } from "@/lib/client-profiles";
import {
  HOLIDAY_STATE_ALL,
  holidayUsesRegion,
  normalizeHolidayDate,
  normalizeHolidayState,
  type HolidayCountryId,
  type HolidayEntry,
  type HolidaysStore,
} from "@/lib/holidays";

export function clientCountryToHolidayId(
  country: ClientCountry | string | undefined,
): HolidayCountryId | null {
  const value = String(country ?? "")
    .trim()
    .toLowerCase();
  if (value === "us" || value === "united states" || value === "unitedstates") {
    return "unitedStates";
  }
  if (value === "ca" || value === "canada") {
    return "canada";
  }
  return null;
}

function toYmd(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

export function formatBusinessDayLabel(date: Date): string {
  const datePart = date.toLocaleDateString("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
  const weekday = date.toLocaleDateString("en-US", { weekday: "long" });
  return `${datePart}, ${weekday}`;
}

function holidayAppliesToClient(
  holiday: HolidayEntry,
  countryId: HolidayCountryId,
  clientState: string,
): boolean {
  if (!holidayUsesRegion(countryId)) {
    return true;
  }
  const scope = normalizeHolidayState(holiday.state, countryId);
  if (scope === HOLIDAY_STATE_ALL) {
    return true;
  }
  const state = clientState.trim().toLowerCase();
  if (!state) {
    return false;
  }
  return scope.toLowerCase() === state;
}

/**
 * Pakistan holidays + client-country holidays (by state/province when set).
 * Used to detect when a scheduled BD lands on a holiday so it can move to the
 * next working day. Weekend-only counting is used for the BD ordinal itself.
 */
export function collectNonBusinessDates(
  store: HolidaysStore,
  clientCountry: ClientCountry | string | undefined,
  clientState: string,
  year: number,
): Set<string> {
  const excluded = new Set<string>();
  const clientHolidayCountryId = clientCountryToHolidayId(clientCountry);

  for (const country of store.countries) {
    const isPakistan = country.id === "pakistan";
    const isClientCountry = country.id === clientHolidayCountryId;
    if (!isPakistan && !isClientCountry) {
      continue;
    }

    for (const holiday of country.holidays) {
      const date = normalizeHolidayDate(holiday.date);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        continue;
      }
      if (!date.startsWith(`${year}-`)) {
        continue;
      }
      if (
        isClientCountry &&
        clientHolidayCountryId &&
        !holidayAppliesToClient(holiday, clientHolidayCountryId, clientState)
      ) {
        continue;
      }
      excluded.add(date);
    }
  }

  return excluded;
}

export function isWeekend(date: Date): boolean {
  const day = date.getDay();
  return day === 0 || day === 6;
}

export function isBusinessDay(date: Date, excludedDates: Set<string>): boolean {
  if (isWeekend(date)) {
    return false;
  }
  return !excludedDates.has(toYmd(date));
}

/**
 * Next working day strictly after `from` (skips weekends + excluded holidays).
 * May cross into a later month.
 */
export function getNextBusinessDay(
  from: Date,
  excludedDates: Set<string>,
): Date {
  const date = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  for (let step = 0; step < 366; step += 1) {
    date.setDate(date.getDate() + 1);
    if (isBusinessDay(date, excludedDates)) {
      return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }
  }
  return new Date(from.getFullYear(), from.getMonth(), from.getDate() + 1);
}

/**
 * If `date` is a weekend or holiday, advance to the next working day.
 * Otherwise return a date-only copy of `date`.
 */
export function ensureWorkingDay(
  date: Date,
  excludedDates: Set<string>,
): Date {
  const day = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  if (isBusinessDay(day, excludedDates)) {
    return day;
  }
  return getNextBusinessDay(day, excludedDates);
}

/** 1-based nth business day within the given month (local calendar). */
export function getNthBusinessDayOfMonth(
  year: number,
  monthIndex: number,
  n: number,
  excludedDates: Set<string>,
): Date | null {
  if (!Number.isFinite(n) || n < 1) {
    return null;
  }

  let count = 0;
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day += 1) {
    const date = new Date(year, monthIndex, day);
    if (!isBusinessDay(date, excludedDates)) {
      continue;
    }
    count += 1;
    if (count === n) {
      return date;
    }
  }
  return null;
}

/**
 * 1-based last nth business day of the month (local calendar).
 * n=1 → last business day; n=2 → second-to-last; skips weekends (and holidays
 * when those are included in excludedDates).
 */
export function getLastNthBusinessDayOfMonth(
  year: number,
  monthIndex: number,
  n: number,
  excludedDates: Set<string>,
): Date | null {
  if (!Number.isFinite(n) || n < 1) {
    return null;
  }

  let count = 0;
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  for (let day = daysInMonth; day >= 1; day -= 1) {
    const date = new Date(year, monthIndex, day);
    if (!isBusinessDay(date, excludedDates)) {
      continue;
    }
    count += 1;
    if (count === n) {
      return date;
    }
  }
  return null;
}

/** Parse schedule ids like bd1 / bd5 / bd10 into the business-day ordinal. */
export function scheduleBusinessDayNumber(scheduleId: string): number | null {
  const match = /^bd(\d+)$/i.exec(scheduleId.trim());
  if (!match) {
    return null;
  }
  const n = Number(match[1]);
  return Number.isFinite(n) && n > 0 ? n : null;
}

/** Parse schedule ids like last-1-bd / last-3-bd / last-3-bd-team into the last-nth ordinal. */
export function scheduleLastBusinessDayNumber(
  scheduleId: string,
): number | null {
  const match = /^last-(\d+)-bd(?:-team)?$/i.exec(scheduleId.trim());
  if (!match) {
    return null;
  }
  const n = Number(match[1]);
  return Number.isFinite(n) && n > 0 ? n : null;
}

/**
 * Resolve a schedule slot (bd5, last-2-bd, …) to a calendar date for the
 * current month.
 *
 * Strategy:
 * 1. Count BD ordinals using weekends only (a holiday on BD3 does not renumber
 *    BD5/BD8).
 * 2. If that date is a client/Pakistan holiday, move that schedule only to the
 *    next working day (skip weekends + holidays). Other BDs stay put.
 * 3. Recalculated each month from `now`.
 */
export function getScheduleBusinessDate(
  scheduleId: string,
  holidays: HolidaysStore,
  clientCountry: ClientCountry | string | undefined,
  clientState: string,
  now: Date = new Date(),
): Date | null {
  const year = now.getFullYear();
  const monthIndex = now.getMonth();
  // Weekend-only count for the BD ordinal (empty holiday set).
  const weekendOnly = new Set<string>();
  // Holidays used only when a scheduled BD lands on a holiday.
  const holidayDates = collectNonBusinessDates(
    holidays,
    clientCountry,
    clientState,
    year,
  );
  // Also include next year's holidays near month-end spillover.
  const nextYearHolidays = collectNonBusinessDates(
    holidays,
    clientCountry,
    clientState,
    year + 1,
  );
  const excludedForShift = new Set<string>([
    ...holidayDates,
    ...nextYearHolidays,
  ]);

  const lastN = scheduleLastBusinessDayNumber(scheduleId);
  let date: Date | null = null;
  if (lastN !== null) {
    date = getLastNthBusinessDayOfMonth(
      year,
      monthIndex,
      lastN,
      weekendOnly,
    );
  } else {
    const n = scheduleBusinessDayNumber(scheduleId);
    if (n === null) {
      return null;
    }
    date = getNthBusinessDayOfMonth(year, monthIndex, n, weekendOnly);
  }

  if (!date) {
    return null;
  }

  // Holiday on this BD → move this schedule to the next working day only.
  return ensureWorkingDay(date, excludedForShift);
}

/** True if any Sat/Sun falls strictly between `from` and `to` (exclusive). */
function hasWeekendBetween(from: Date, to: Date): boolean {
  const cursor = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  const end = new Date(to.getFullYear(), to.getMonth(), to.getDate());
  cursor.setDate(cursor.getDate() + 1);
  while (cursor < end) {
    if (isWeekend(cursor)) {
      return true;
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return false;
}

/**
 * BD7 email deadline: BD8 by default; if a weekend falls between BD7 and BD8
 * (e.g. BD7 Fri → BD8 Mon), use BD7 instead.
 */
export function getBd7DeadlineDate(
  holidays: HolidaysStore,
  clientCountry: ClientCountry | string | undefined,
  clientState: string,
  now: Date = new Date(),
): Date | null {
  const bd7 = getScheduleBusinessDate(
    "bd7",
    holidays,
    clientCountry,
    clientState,
    now,
  );
  const bd8 = getScheduleBusinessDate(
    "bd8",
    holidays,
    clientCountry,
    clientState,
    now,
  );
  if (!bd7) {
    return bd8;
  }
  if (!bd8) {
    return bd7;
  }
  if (hasWeekendBetween(bd7, bd8)) {
    return bd7;
  }
  return bd8;
}
