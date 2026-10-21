import type { HolidayCountryId } from "@/lib/holidays";

export type HolidayKind = "fixed" | "floating" | "lunar";

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

function toYmd(date: Date): string {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function normalizeName(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, " ");
}

/** US OPM-style: Sat → Fri, Sun → Mon. */
export function applyWeekendObserved(date: Date): Date {
  const next = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = next.getDay();
  if (day === 6) {
    next.setDate(next.getDate() - 1);
  } else if (day === 0) {
    next.setDate(next.getDate() + 1);
  }
  return next;
}

/** n = 1 → first, n = 2 → second, … weekday: 0=Sun … 6=Sat */
export function nthWeekdayOfMonth(
  year: number,
  monthIndex: number,
  weekday: number,
  n: number,
): Date {
  const first = new Date(year, monthIndex, 1);
  const offset = (weekday - first.getDay() + 7) % 7;
  const day = 1 + offset + (n - 1) * 7;
  return new Date(year, monthIndex, day);
}

export function lastWeekdayOfMonth(
  year: number,
  monthIndex: number,
  weekday: number,
): Date {
  const last = new Date(year, monthIndex + 1, 0);
  const offset = (last.getDay() - weekday + 7) % 7;
  return new Date(year, monthIndex, last.getDate() - offset);
}

/** Anonymous Gregorian algorithm → Easter Sunday (Western). */
export function easterSunday(year: number): Date {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

/** Monday preceding May 25 (Victoria Day). */
export function mondayBeforeMay25(year: number): Date {
  const may25 = new Date(year, 4, 25);
  const day = may25.getDay();
  const back = day === 1 ? 7 : (day + 6) % 7;
  return new Date(year, 4, 25 - back);
}

export function classifyHolidayName(
  name: string,
  countryId?: HolidayCountryId,
): HolidayKind {
  const n = normalizeName(name);

  if (
    n.includes("eid") ||
    n.includes("arafa") ||
    n.includes("ashura") ||
    n.includes("moharram") ||
    n.includes("muharram") ||
    n.includes("rabi")
  ) {
    return "lunar";
  }

  // Pakistan Labour Day is fixed May 1; US/Canada Labour/Labor Day floats.
  if (n.includes("labor day") || n.includes("labour day")) {
    return countryId === "pakistan" ? "fixed" : "floating";
  }

  if (
    n.includes("martin luther") ||
    n.includes("washington") ||
    n.includes("presidents") ||
    n.includes("memorial day") ||
    n.includes("columbus") ||
    n.includes("thanksgiving") ||
    n.includes("good friday") ||
    n.includes("victoria day")
  ) {
    return "floating";
  }

  return "fixed";
}

/**
 * Compute the observance date for a known holiday in `year`.
 * Returns null for lunar / unknown names (caller keeps year-bump).
 */
export function computeHolidayDate(
  name: string,
  countryId: HolidayCountryId,
  year: number,
): string | null {
  const n = normalizeName(name);
  const kind = classifyHolidayName(name, countryId);
  if (kind === "lunar") {
    return null;
  }

  // —— Floating ——
  if (n.includes("martin luther")) {
    return toYmd(nthWeekdayOfMonth(year, 0, 1, 3));
  }
  if (n.includes("washington") || n.includes("presidents")) {
    return toYmd(nthWeekdayOfMonth(year, 1, 1, 3));
  }
  if (n.includes("memorial day")) {
    return toYmd(lastWeekdayOfMonth(year, 4, 1));
  }
  if (n.includes("labor day") || n.includes("labour day")) {
    if (countryId === "pakistan") {
      return toYmd(new Date(year, 4, 1));
    }
    return toYmd(nthWeekdayOfMonth(year, 8, 1, 1));
  }
  if (n.includes("columbus")) {
    return toYmd(nthWeekdayOfMonth(year, 9, 1, 2));
  }
  if (n.includes("thanksgiving")) {
    if (countryId === "canada") {
      return toYmd(nthWeekdayOfMonth(year, 9, 1, 2));
    }
    return toYmd(nthWeekdayOfMonth(year, 10, 4, 4));
  }
  if (n.includes("good friday")) {
    const easter = easterSunday(year);
    easter.setDate(easter.getDate() - 2);
    return toYmd(easter);
  }
  if (n.includes("victoria day")) {
    return toYmd(mondayBeforeMay25(year));
  }

  // —— Fixed (with weekend observed where OPM-style) ——
  if (n.includes("new year")) {
    const d = new Date(year, 0, 1);
    return toYmd(
      countryId === "unitedStates" ? applyWeekendObserved(d) : d,
    );
  }
  if (n.includes("juneteenth")) {
    return toYmd(applyWeekendObserved(new Date(year, 5, 19)));
  }
  if (n.includes("independence day") && countryId === "unitedStates") {
    return toYmd(applyWeekendObserved(new Date(year, 6, 4)));
  }
  if (n.includes("canada day")) {
    const d = new Date(year, 6, 1);
    // Common practice: if Sunday → Monday
    if (d.getDay() === 0) {
      d.setDate(d.getDate() + 1);
    }
    return toYmd(d);
  }
  if (n.includes("veterans day") || n.includes("remembrance day")) {
    const d = new Date(year, 10, 11);
    return toYmd(
      countryId === "unitedStates" ? applyWeekendObserved(d) : d,
    );
  }
  if (n.includes("christmas") && !n.includes("boxing")) {
    const d = new Date(year, 11, 25);
    return toYmd(
      countryId === "unitedStates" ? applyWeekendObserved(d) : d,
    );
  }
  if (n.includes("boxing day")) {
    return toYmd(new Date(year, 11, 26));
  }
  if (n.includes("truth and reconciliation")) {
    return toYmd(new Date(year, 8, 30));
  }
  if (n.includes("kashmir day")) {
    return toYmd(new Date(year, 1, 5));
  }
  if (n.includes("pakistan day")) {
    return toYmd(new Date(year, 2, 23));
  }
  if (n.includes("independence day") && countryId === "pakistan") {
    return toYmd(new Date(year, 7, 14));
  }
  if (n.includes("quaid")) {
    return toYmd(new Date(year, 11, 25));
  }

  return null;
}

export function isLunarHolidayName(name: string): boolean {
  return classifyHolidayName(name) === "lunar";
}

export function isFloatingHolidayName(
  name: string,
  countryId?: HolidayCountryId,
): boolean {
  return classifyHolidayName(name, countryId) === "floating";
}
