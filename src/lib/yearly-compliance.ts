/** Yearly Compliance → Yearly Compliance Timeline */
export const YEARLY_COMPLIANCE_TIMELINES: Record<string, string> = {
  "1120": "30th April",
  "1040": "30th April",
  "1099": "31st January",
  "1120S": "15th March",
  "1065": "15th March",
  T1: "31st March",
  "T2 (Calendar Year, 31 Dec)": "31st March",
  "T2 (Fiscal Year, 30 June)": "30th Sep",
  T4: "31st January",
  T5: "31st January",
  NIL: "NIL",
  "990 Form": "15th Nov",
};

/** US clients: options that start with a digit (1120, 1120S, …) plus NIL. */
export function isUsYearlyComplianceOption(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) {
    return false;
  }
  if (trimmed.toLowerCase() === "nil") {
    return true;
  }
  return /^\d/.test(trimmed);
}

/** Canadian clients: options that start with a letter (T1, T2, …) plus NIL. */
export function isCaYearlyComplianceOption(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) {
    return false;
  }
  if (trimmed.toLowerCase() === "nil") {
    return true;
  }
  return /^[A-Za-z]/.test(trimmed);
}

/**
 * Filter Yearly Compliance dropdown by Client Profile country.
 * US → numeric-leading + NIL; CA → letter-leading + NIL;
 * Not Sure → full list; no country yet → US-style.
 */
export function filterYearlyComplianceOptionsForCountry(
  options: string[],
  country: string | undefined | null,
): string[] {
  const normalized = String(country ?? "")
    .trim()
    .toLowerCase();
  if (
    normalized === "ca" ||
    normalized === "canada" ||
    normalized === "canadian"
  ) {
    return options.filter(isCaYearlyComplianceOption);
  }
  if (normalized === "not sure" || normalized === "ns") {
    return options;
  }
  // US (and default when country is not set yet)
  return options.filter(isUsYearlyComplianceOption);
}

export const YEARLY_COMPLIANCE_OPTIONS = Object.keys(
  YEARLY_COMPLIANCE_TIMELINES,
);

function normalizeKey(value: string): string {
  return value.trim().toLowerCase();
}

const TIMELINE_BY_NORMALIZED = new Map(
  Object.entries(YEARLY_COMPLIANCE_TIMELINES).map(([key, timeline]) => [
    normalizeKey(key),
    timeline,
  ]),
);

/** Legacy labels that map to the new T2 options when resolving timelines. */
const LEGACY_ALIASES: Record<string, string> = {
  t2: "T2 (Calendar Year, 31 Dec)",
};

export function getYearlyComplianceTimeline(compliance: string): string {
  const key = normalizeKey(compliance);
  const direct = TIMELINE_BY_NORMALIZED.get(key);
  if (direct) {
    return direct;
  }

  const alias = LEGACY_ALIASES[key];
  if (alias) {
    return YEARLY_COMPLIANCE_TIMELINES[alias] ?? "";
  }

  return "";
}

export function resolveYearlyComplianceTimelines(
  selected: string[],
): string {
  const onlyNil =
    selected.length === 1 && selected[0].trim().toLowerCase() === "nil";
  if (onlyNil) {
    return "NIL";
  }

  return selected
    .map((item) => {
      const compliance = item.trim();
      if (!compliance) {
        return "";
      }
      const timeline = getYearlyComplianceTimeline(compliance);
      return timeline ? `${compliance} - ${timeline}` : "";
    })
    .filter(Boolean)
    .join("\n");
}
