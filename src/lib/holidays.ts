import { CANADIAN_PROVINCES, US_STATES } from "@/lib/regions";

export type HolidayCountryId = "pakistan" | "unitedStates" | "canada";

/** Federal / nationwide scope for US states or Canadian provinces. */
export const HOLIDAY_STATE_ALL = "All";

export type HolidayEntry = {
  id: string;
  name: string;
  /** Stored as YYYY-MM-DD */
  date: string;
  /**
   * US / Canada: "All" for federal holidays, or a specific state / province.
   * Omitted / unused for other countries.
   */
  state?: string;
};

export type HolidayCountry = {
  id: HolidayCountryId;
  name: string;
  holidays: HolidayEntry[];
};

export type HolidaysStore = {
  countries: HolidayCountry[];
};

export const HOLIDAY_COUNTRIES: {
  id: HolidayCountryId;
  name: string;
}[] = [
  { id: "pakistan", name: "Pakistan" },
  { id: "unitedStates", name: "United States" },
  { id: "canada", name: "Canada" },
];

function makeId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function holidayUsesRegion(countryId: HolidayCountryId): boolean {
  return countryId === "unitedStates" || countryId === "canada";
}

export function getHolidayRegionLabel(countryId: HolidayCountryId): string {
  return countryId === "canada" ? "Province" : "States";
}

export function getHolidayRegionOptions(countryId: HolidayCountryId): string[] {
  if (countryId === "canada") {
    return [HOLIDAY_STATE_ALL, ...CANADIAN_PROVINCES];
  }
  if (countryId === "unitedStates") {
    return [HOLIDAY_STATE_ALL, ...US_STATES];
  }
  return [HOLIDAY_STATE_ALL];
}

export function getUsHolidayStateOptions(): string[] {
  return getHolidayRegionOptions("unitedStates");
}

export function normalizeHolidayState(
  value: unknown,
  countryId: HolidayCountryId = "unitedStates",
): string {
  const trimmed = String(value ?? "").trim();
  if (!trimmed || trimmed.toLowerCase() === "all") {
    return HOLIDAY_STATE_ALL;
  }
  const options =
    countryId === "canada"
      ? CANADIAN_PROVINCES
      : countryId === "unitedStates"
        ? US_STATES
        : [];
  const match = options.find(
    (region) => region.toLowerCase() === trimmed.toLowerCase(),
  );
  return match ?? HOLIDAY_STATE_ALL;
}

export function createHolidayEntry(
  name: string,
  date: string,
  state?: string,
  countryId: HolidayCountryId = "unitedStates",
): HolidayEntry {
  const entry: HolidayEntry = {
    id: makeId("holiday"),
    name: name.trim(),
    date: normalizeHolidayDate(date),
  };
  if (state !== undefined) {
    entry.state = normalizeHolidayState(state, countryId);
  }
  return entry;
}

/** Accept YYYY-MM-DD or DD-MM-YYYY; store YYYY-MM-DD. */
export function normalizeHolidayDate(value: string): string {
  const trimmed = value.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    return trimmed;
  }
  const match = trimmed.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (!match) {
    return trimmed;
  }
  return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
}

/** Display YYYY-MM-DD as DD-MM-YYYY. */
export function formatHolidayDate(value: string): string {
  const match = value.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return value.trim();
  }
  return `${match[3]}-${match[2]}-${match[1]}`;
}

export function getHolidayWeekday(date: string): string {
  const normalized = normalizeHolidayDate(date);
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return "—";
  }
  const parsed = new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
  );
  if (Number.isNaN(parsed.getTime())) {
    return "—";
  }
  return parsed.toLocaleDateString("en-US", { weekday: "long" });
}

const PAKISTAN_2026_SEED: { name: string; date: string }[] = [
  { name: "Kashmir Day", date: "2026-02-05" },
  { name: "Pakistan Day", date: "2026-03-23" },
  { name: "Eid-ul-Fitr - Day 1", date: "2026-03-21" },
  { name: "Eid-ul-Fitr - Day 2", date: "2026-03-22" },
  { name: "Eid-ul-Fitr - Day 3", date: "2026-03-23" },
  { name: "Labour Day", date: "2026-05-01" },
  { name: "Day of Arafa", date: "2026-05-26" },
  { name: "Eid ul-Adha - Day 1", date: "2026-05-27" },
  { name: "Eid ul-Adha - Day 2", date: "2026-05-28" },
  { name: "Ashura - 9th Moharram", date: "2026-06-24" },
  { name: "Ashura - 10th Moharram", date: "2026-06-25" },
  { name: "Independence Day", date: "2026-08-14" },
  { name: "12th of Rabi-ul-Awal", date: "2026-08-25" },
  { name: "Quaid-e-Azam Day", date: "2026-12-25" },
];

/** OPM federal holiday schedule for 2026 (observed dates). */
const UNITED_STATES_2026_SEED: { name: string; date: string }[] = [
  { name: "New Year's Day", date: "2026-01-01" },
  { name: "Birthday of Martin Luther King, Jr.", date: "2026-01-19" },
  { name: "Washington's Birthday", date: "2026-02-16" },
  { name: "Memorial Day", date: "2026-05-25" },
  { name: "Juneteenth National Independence Day", date: "2026-06-19" },
  { name: "Independence Day", date: "2026-07-03" },
  { name: "Labor Day", date: "2026-09-07" },
  { name: "Columbus Day", date: "2026-10-12" },
  { name: "Veterans Day", date: "2026-11-11" },
  { name: "Thanksgiving Day", date: "2026-11-26" },
  { name: "Christmas Day", date: "2026-12-25" },
];

/** Canada Labour Code general holidays for 2026. */
const CANADA_2026_SEED: { name: string; date: string }[] = [
  { name: "New Year's Day", date: "2026-01-01" },
  { name: "Good Friday", date: "2026-04-03" },
  { name: "Victoria Day", date: "2026-05-18" },
  { name: "Canada Day", date: "2026-07-01" },
  { name: "Labour Day", date: "2026-09-07" },
  { name: "National Day for Truth and Reconciliation", date: "2026-09-30" },
  { name: "Thanksgiving Day", date: "2026-10-12" },
  { name: "Remembrance Day", date: "2026-11-11" },
  { name: "Christmas Day", date: "2026-12-25" },
  { name: "Boxing Day", date: "2026-12-26" },
];

const COUNTRY_SEEDS: Partial<
  Record<HolidayCountryId, { name: string; date: string }[]>
> = {
  pakistan: PAKISTAN_2026_SEED,
  unitedStates: UNITED_STATES_2026_SEED,
  canada: CANADA_2026_SEED,
};

function seedHolidays(
  countryId: HolidayCountryId,
  entries: { name: string; date: string }[],
): HolidayEntry[] {
  const prefix = countryId === "unitedStates" ? "us" : countryId.slice(0, 2);
  return entries.map((entry, index) => {
    const holiday: HolidayEntry = {
      id: `seed-${prefix}-${index + 1}-${entry.name.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`,
      name: entry.name,
      date: entry.date,
    };
    if (holidayUsesRegion(countryId)) {
      holiday.state = HOLIDAY_STATE_ALL;
    }
    return holiday;
  });
}

export function createDefaultHolidays(): HolidaysStore {
  return {
    countries: HOLIDAY_COUNTRIES.map((country) => ({
      id: country.id,
      name: country.name,
      holidays: COUNTRY_SEEDS[country.id]
        ? seedHolidays(country.id, COUNTRY_SEEDS[country.id]!)
        : [],
    })),
  };
}

export function normalizeHolidays(value: unknown): HolidaysStore {
  const defaults = createDefaultHolidays();
  const incoming =
    typeof value === "object" && value !== null
      ? (value as { countries?: unknown })
      : {};
  const incomingCountries = Array.isArray(incoming.countries)
    ? incoming.countries
    : [];

  const byId = new Map<HolidayCountryId, HolidayEntry[]>();

  for (const country of incomingCountries) {
    if (typeof country !== "object" || country === null) {
      continue;
    }
    const raw = country as { id?: unknown; holidays?: unknown };
    if (
      raw.id !== "pakistan" &&
      raw.id !== "unitedStates" &&
      raw.id !== "canada"
    ) {
      continue;
    }
    const countryId: HolidayCountryId = raw.id;

    const holidays = Array.isArray(raw.holidays)
      ? raw.holidays
          .map((item) => {
            if (typeof item !== "object" || item === null) {
              return null;
            }
            const holiday = item as Partial<HolidayEntry>;
            const name = String(holiday.name ?? "").trim();
            const date = normalizeHolidayDate(String(holiday.date ?? ""));
            if (!name || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
              return null;
            }
            const entry: HolidayEntry = {
              id:
                typeof holiday.id === "string" && holiday.id
                  ? holiday.id
                  : makeId("holiday"),
              name,
              date,
            };
            if (holidayUsesRegion(countryId)) {
              entry.state = normalizeHolidayState(holiday.state, countryId);
            }
            return entry;
          })
          .filter((item): item is HolidayEntry => item !== null)
      : [];

    byId.set(countryId, holidays);
  }

  return {
    countries: defaults.countries.map((country) => {
      if (byId.has(country.id)) {
        const incomingHolidays = byId.get(country.id) ?? [];
        // Fill seeded defaults when a known country list is still empty.
        if (incomingHolidays.length === 0 && country.holidays.length > 0) {
          return country;
        }
        return {
          ...country,
          holidays: incomingHolidays,
        };
      }
      return country;
    }),
  };
}
