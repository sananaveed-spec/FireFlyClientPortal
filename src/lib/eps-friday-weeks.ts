/**
 * EPS-only: Week 1 Friday, Week 2 Friday, … for every Friday in the month.
 * One Active toggle (`eps-fridays`) enables all week slots.
 */

export const EPS_FRIDAYS_TOGGLE_ID = "eps-fridays";

const EPS_WEEK_FRIDAY_RE = /^eps-week-(\d+)-friday$/i;

export type EpsFridayWeekItem = {
  id: string;
  label: string;
  /** 1-based week index within the month. */
  weekNumber: number;
  date: Date;
};

export function isEpsFridaysToggleId(value: string): boolean {
  return value.trim().toLowerCase() === EPS_FRIDAYS_TOGGLE_ID;
}

export function isEpsFridayWeekScheduleId(value: string): boolean {
  return EPS_WEEK_FRIDAY_RE.test(value.trim());
}

export function parseEpsFridayWeekNumber(scheduleId: string): number | null {
  const match = EPS_WEEK_FRIDAY_RE.exec(scheduleId.trim());
  if (!match) {
    return null;
  }
  const n = Number(match[1]);
  return Number.isFinite(n) && n >= 1 ? n : null;
}

export function epsFridayWeekScheduleId(weekNumber: number): string {
  return `eps-week-${weekNumber}-friday`;
}

/** Every Friday in the calendar month (local Y/M/D dates). */
export function getFridaysOfMonth(year: number, monthIndex: number): Date[] {
  const fridays: Date[] = [];
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day += 1) {
    const date = new Date(year, monthIndex, day);
    if (date.getDay() === 5) {
      fridays.push(date);
    }
  }
  return fridays;
}

export function getEpsFridayWeekItemsForMonth(
  now: Date = new Date(),
): EpsFridayWeekItem[] {
  const year = now.getFullYear();
  const monthIndex = now.getMonth();
  return getFridaysOfMonth(year, monthIndex).map((date, index) => {
    const weekNumber = index + 1;
    return {
      id: epsFridayWeekScheduleId(weekNumber),
      label: `Week ${weekNumber} Friday`,
      weekNumber,
      date,
    };
  });
}

/** This month’s Friday for `eps-week-N-friday`, or null if that week doesn’t exist. */
export function getEpsFridayDateForSchedule(
  scheduleId: string,
  now: Date = new Date(),
): Date | null {
  const weekNumber = parseEpsFridayWeekNumber(scheduleId);
  if (weekNumber === null) {
    return null;
  }
  const fridays = getFridaysOfMonth(now.getFullYear(), now.getMonth());
  return fridays[weekNumber - 1] ?? null;
}

/**
 * Active check for an EPS week row: driven by the single `eps-fridays` toggle.
 */
export function isEpsFridaysActive(
  activeScheduleIds: readonly string[],
): boolean {
  return activeScheduleIds.some((id) => isEpsFridaysToggleId(id));
}
