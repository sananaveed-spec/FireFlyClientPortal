const MONTH_LABELS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
] as const;

export const YEAR_END_MONTHS = MONTH_LABELS.map((label, index) => ({
  value: String(index + 1).padStart(2, "0"),
  label,
}));

export const YEAR_END_DAYS = Array.from({ length: 31 }, (_, index) => {
  const day = index + 1;
  return { value: String(day), label: String(day) };
});

const MONTH_BY_LABEL = new Map(
  MONTH_LABELS.map((label, index) => [label.toLowerCase(), index + 1]),
);

export type YearEndParts = {
  day: string;
  month: string;
};

/** Format as "31 Dec" (day + month only, no year). */
export function formatYearEnd(day: number, month: number): string {
  if (day < 1 || day > 31 || month < 1 || month > 12) {
    return "";
  }
  return `${day} ${MONTH_LABELS[month - 1]}`;
}

export function parseYearEnd(value: string): YearEndParts | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  // Legacy full date: YYYY-MM-DD
  const iso = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) {
    return {
      day: String(Number.parseInt(iso[3], 10)),
      month: iso[2],
    };
  }

  // MM-DD
  const md = trimmed.match(/^(\d{1,2})-(\d{1,2})$/);
  if (md) {
    return {
      day: String(Number.parseInt(md[2], 10)),
      month: String(Number.parseInt(md[1], 10)).padStart(2, "0"),
    };
  }

  // "31 Dec" / "31 December"
  const named = trimmed.match(/^(\d{1,2})\s+([A-Za-z]+)$/);
  if (named) {
    const monthKey = named[2].toLowerCase().slice(0, 3);
    const monthNum = MONTH_BY_LABEL.get(monthKey);
    if (!monthNum) {
      return null;
    }
    return {
      day: String(Number.parseInt(named[1], 10)),
      month: String(monthNum).padStart(2, "0"),
    };
  }

  return null;
}

/** Normalize stored value to "31 Dec" (or empty). */
export function normalizeYearEnd(value: string): string {
  const parts = parseYearEnd(value);
  if (!parts) {
    return value.trim();
  }
  return formatYearEnd(
    Number.parseInt(parts.day, 10),
    Number.parseInt(parts.month, 10),
  );
}
