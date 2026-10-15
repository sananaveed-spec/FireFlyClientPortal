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

/** e.g. October 2026 → { yearFolder: "2026", monthFolder: "10.Oct26" } */
export function sharePointPeriodFolder(date: Date = new Date()): {
  yearFolder: string;
  monthFolder: string;
} {
  const year = date.getFullYear();
  const yy = String(year).slice(-2);
  const monthIndex = date.getMonth();
  const n = String(monthIndex + 1).padStart(2, "0");
  return {
    yearFolder: String(year),
    monthFolder: `${n}.${MONTH_LABELS[monthIndex]}${yy}`,
  };
}

/** All twelve month folder names for a calendar year (01.JanYY … 12.DecYY). */
export function sharePointYearMonthFolderNames(year: number): {
  yearFolder: string;
  monthFolders: string[];
} {
  const yy = String(year).slice(-2);
  return {
    yearFolder: String(year),
    monthFolders: MONTH_LABELS.map((label, index) => {
      const n = String(index + 1).padStart(2, "0");
      return `${n}.${label}${yy}`;
    }),
  };
}
