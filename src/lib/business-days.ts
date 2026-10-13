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

/** When true, business-day math skips weekends only (holidays are not deducted). */
const IGNORE_HOLIDAYS_IN_BUSINESS_DAYS = true;

/** Pakistan holidays + client-country holidays (by state/province when set). */
export function collectNonBusinessDates(
  store: HolidaysStore,
  clientCountry: ClientCountry | string | undefined,
  clientState: string,
  year: number,
): Set<string> {
  if (IGNORE_HOLIDAYS_IN_BUSINESS_DAYS) {
    return new Set();
  }

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

/** Parse schedule ids like last-1-bd / last-3-bd into the last-nth ordinal. */
export function scheduleLastBusinessDayNumber(
  scheduleId: string,
): number | null {
  const match = /^last-(\d+)-bd$/i.exec(scheduleId.trim());
  if (!match) {
    return null;
  }
  const n = Number(match[1]);
  return Number.isFinite(n) && n > 0 ? n : null;
}

export function getScheduleBusinessDate(
  scheduleId: string,
  holidays: HolidaysStore,
  clientCountry: ClientCountry | string | undefined,
  clientState: string,
  now: Date = new Date(),
): Date | null {
  const year = now.getFullYear();
  const monthIndex = now.getMonth();
  const excluded = collectNonBusinessDates(
    holidays,
    clientCountry,
    clientState,
    year,
  );

  const lastN = scheduleLastBusinessDayNumber(scheduleId);
  if (lastN !== null) {
    return getLastNthBusinessDayOfMonth(year, monthIndex, lastN, excluded);
  }

  const n = scheduleBusinessDayNumber(scheduleId);
  if (n === null) {
    return null;
  }
  return getNthBusinessDayOfMonth(year, monthIndex, n, excluded);
}
