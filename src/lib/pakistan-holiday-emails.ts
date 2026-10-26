import { applyPeriodPlaceholder } from "@/lib/email-templates";
import { formatHolidayDate, type HolidayEntry } from "@/lib/holidays";
import { isValidEmail, normalizeEmail } from "@/lib/portal-users";
import {
  CLIENT_SEND_HOUR,
  resolveClientTimeZone,
  USER_TIME_ZONE,
  zonedWallTimeToUtc,
} from "@/lib/time-zones";

export type PakistanHolidayEmailStatus = "scheduled" | "sent";

/** Team notice (day before) vs client notice (7 business days before). */
export type PakistanHolidayEmailAudience = "team" | "client";

export type PakistanHolidayEmailEntry = {
  /** Stable series key, e.g. "2026:eid-ul-fitr". */
  groupKey: string;
  audience: PakistanHolidayEmailAudience;
  /** First day of the series (YYYY-MM-DD). */
  holidayDate: string;
  /** Every calendar date covered by this one email. */
  holidayDates: string[];
  /** Holiday ids covered by this one email. */
  holidayIds: string[];
  holidayName: string;
  status: PakistanHolidayEmailStatus;
  /** When the email is/was scheduled to send (ISO). */
  scheduledFor: string;
  sentAt: string;
};

export type PakistanHolidayEmailsStore = {
  entries: PakistanHolidayEmailEntry[];
};

export type PakistanHolidayEmailGroup = {
  groupKey: string;
  /** Display name for the series (e.g. "Eid-ul-Fitr"). */
  seriesName: string;
  holidays: HolidayEntry[];
};

export type PakistanHolidayEmailPlan = {
  groupKey: string;
  audience: PakistanHolidayEmailAudience;
  holidayDate: string;
  holidayDates: string[];
  holidayIds: string[];
  holidayName: string;
  subject: string;
  body: string;
  /** ISO send time; omit / past → client should send immediately. */
  sendAt: string | null;
  /** Mark as sent (immediate) vs scheduled after Graph accepts. */
  markAs: PakistanHolidayEmailStatus;
  /** Client plans: recipients for this timezone batch. */
  recipients?: string[];
  /** IANA zone used for sendAt (client plans). */
  timeZone?: string;
};

export type PakistanHolidayClientRecipientGroup = {
  timeZone: string;
  emails: string[];
};

export type PakistanHolidayEmailStatusLabel = "Scheduled" | "Sent" | "—";

export const PAKISTAN_HOLIDAY_EMAIL_TEMPLATE_ID = "template-16";
export const PAKISTAN_HOLIDAY_CLIENT_EMAIL_TEMPLATE_ID = "template-17";

/** Business days before the holiday for client notices (Sat/Sun skipped). */
export const PAKISTAN_HOLIDAY_CLIENT_EMAIL_LEAD_BUSINESS_DAYS = 7;

/** Preferred From/To mailbox when scheduling Pakistan holiday emails. */
export const PAKISTAN_HOLIDAY_EMAIL_SENDER =
  "ayesha@fireflyaccountancy.com";

/** Portal + Microsoft accounts allowed to schedule/send Pakistan holiday emails. */
export const PAKISTAN_HOLIDAY_EMAIL_SENDERS = [
  PAKISTAN_HOLIDAY_EMAIL_SENDER,
  "sana@epsfresno.com",
] as const;

/**
 * Users who should always see the Upcoming holiday emails preview / status
 * (in addition to any signed-in portal user). Listed in UI copy.
 */
export const PAKISTAN_HOLIDAY_EMAIL_VIEWERS = [
  "sana@epsfresno.com",
  "sana@fireflyaccountancy.com",
] as const;

export function isPakistanHolidayEmailSender(email: string): boolean {
  const normalized = email.trim().toLowerCase();
  return (PAKISTAN_HOLIDAY_EMAIL_SENDERS as readonly string[]).includes(
    normalized,
  );
}

export function pakistanHolidayEmailSendersLabel(): string {
  return PAKISTAN_HOLIDAY_EMAIL_SENDERS.join(" or ");
}

export function createEmptyPakistanHolidayEmailsStore(): PakistanHolidayEmailsStore {
  return { entries: [] };
}

/**
 * Series key for grouping multi-day holidays into one email.
 * "Eid-ul-Fitr - Day 1/2/3" → "eid-ul-fitr"
 * "Ashura - 9th Moharram" → "ashura"
 * "Kashmir Day" → "kashmir day"
 */
export function pakistanHolidaySeriesKey(name: string): string {
  const trimmed = name.trim();
  const dayStripped = trimmed.replace(/\s*-\s*Day\s+\d+\s*$/i, "").trim();
  if (dayStripped !== trimmed && dayStripped) {
    return dayStripped.toLowerCase();
  }
  const ordinalStripped = trimmed
    .replace(/\s*-\s*\d+(?:st|nd|rd|th)\b.*$/i, "")
    .trim();
  if (ordinalStripped !== trimmed && ordinalStripped) {
    return ordinalStripped.toLowerCase();
  }
  return trimmed.toLowerCase();
}

function slugSeriesKey(seriesKey: string): string {
  return seriesKey
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

export function pakistanHolidayGroupKey(
  seriesKey: string,
  firstDate: string,
): string {
  const year = firstDate.slice(0, 4) || "0000";
  return `${year}:${slugSeriesKey(seriesKey) || "holiday"}`;
}

function holidayDayNumber(name: string): number | null {
  const match = name.trim().match(/\bDay\s+(\d+)\s*$/i);
  return match ? Number(match[1]) : null;
}

function holidayBaseLabel(name: string): string {
  const trimmed = name.trim();
  const dayStripped = trimmed.replace(/\s*-\s*Day\s+\d+\s*$/i, "").trim();
  return dayStripped || trimmed;
}

/** Next calendar day for a YYYY-MM-DD value (UTC date parts). */
function nextCalendarDate(date: string): string | null {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return null;
  }
  const [year, month, day] = date.split("-").map(Number);
  const utc = new Date(Date.UTC(year!, month! - 1, day!));
  utc.setUTCDate(utc.getUTCDate() + 1);
  const y = utc.getUTCFullYear();
  const m = String(utc.getUTCMonth() + 1).padStart(2, "0");
  const d = String(utc.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function seriesDisplayName(holidays: HolidayEntry[]): string {
  if (holidays.length === 0) {
    return "";
  }
  if (holidays.length === 1) {
    return holidays[0]!.name;
  }
  const days = holidays.map((item) => holidayDayNumber(item.name));
  const allNumbered = days.every((n) => n !== null);
  const bases = [
    ...new Set(holidays.map((item) => holidayBaseLabel(item.name))),
  ];
  const seriesKeys = [
    ...new Set(holidays.map((item) => pakistanHolidaySeriesKey(item.name))),
  ];
  const baseLabel = bases.join(" / ");
  if (allNumbered) {
    const nums = days as number[];
    const min = Math.min(...nums);
    const max = Math.max(...nums);
    return min === max
      ? `${baseLabel} (Day ${min})`
      : `${baseLabel} (Day ${min}–${max})`;
  }
  if (seriesKeys.length === 1) {
    return seriesKeys[0]!.replace(/\b\w/g, (ch) => ch.toUpperCase());
  }
  if (bases.length === 1) {
    return bases[0]!;
  }
  return baseLabel;
}

function groupKeyForHolidays(holidays: HolidayEntry[]): string {
  const firstDate = holidays[0]!.date;
  const year = firstDate.slice(0, 4) || "0000";
  const bases = [
    ...new Set(
      holidays.map((item) => slugSeriesKey(pakistanHolidaySeriesKey(item.name))),
    ),
  ].filter(Boolean);
  return `${year}:${bases.join("+") || "holiday"}`;
}

/**
 * Same series name, or consecutive calendar days with sequential Day 1/2/3…
 * (e.g. Day of Arafa - Day 1 + Eid ul-Adha - Day 2/3 → one email).
 */
function shouldMergeHolidayIntoRun(
  run: HolidayEntry[],
  holiday: HolidayEntry,
): boolean {
  const last = run[run.length - 1]!;
  const sameSeries =
    pakistanHolidaySeriesKey(last.name) ===
    pakistanHolidaySeriesKey(holiday.name);

  if (holiday.date === last.date) {
    return sameSeries;
  }

  const isNextDay = nextCalendarDate(last.date) === holiday.date;
  if (!isNextDay) {
    return sameSeries;
  }

  if (sameSeries) {
    return true;
  }

  const lastDay = holidayDayNumber(last.name);
  const nextDay = holidayDayNumber(holiday.name);
  return (
    lastDay !== null && nextDay !== null && nextDay === lastDay + 1
  );
}

/**
 * One group = one team email. Multi-day series (Eid Day 1–3, Arafa+Eid, etc.)
 * share a group.
 */
export function groupPakistanHolidaysForEmail(
  holidays: HolidayEntry[],
): PakistanHolidayEmailGroup[] {
  const sorted = [...holidays].sort((a, b) => {
    const byDate = a.date.localeCompare(b.date);
    if (byDate !== 0) {
      return byDate;
    }
    return a.name.localeCompare(b.name);
  });

  const runs: HolidayEntry[][] = [];
  for (const holiday of sorted) {
    const current = runs[runs.length - 1];
    if (current && shouldMergeHolidayIntoRun(current, holiday)) {
      current.push(holiday);
    } else {
      runs.push([holiday]);
    }
  }

  // Merge non-adjacent same-series leftovers (same year) that the day-walk split.
  const bySeries = new Map<string, HolidayEntry[]>();
  const sequentialDayRuns: HolidayEntry[][] = [];
  for (const run of runs) {
    const keys = new Set(
      run.map((item) => pakistanHolidaySeriesKey(item.name)),
    );
    const numbered = run.every((item) => holidayDayNumber(item.name) !== null);
    const mixedBases = keys.size > 1;
    if (numbered && mixedBases) {
      sequentialDayRuns.push(run);
      continue;
    }
    const seriesKey = pakistanHolidaySeriesKey(run[0]!.name);
    const year = run[0]!.date.slice(0, 4) || "0000";
    const mapKey = `${year}:${seriesKey}`;
    const list = bySeries.get(mapKey) ?? [];
    list.push(...run);
    bySeries.set(mapKey, list);
  }

  const groups: PakistanHolidayEmailGroup[] = [
    ...sequentialDayRuns.map((ordered) => {
      const holidaysInGroup = [...ordered].sort((a, b) =>
        a.date.localeCompare(b.date),
      );
      return {
        groupKey: groupKeyForHolidays(holidaysInGroup),
        seriesName: seriesDisplayName(holidaysInGroup),
        holidays: holidaysInGroup,
      };
    }),
    ...[...bySeries.values()].map((list) => {
      const ordered = [...list].sort((a, b) => a.date.localeCompare(b.date));
      return {
        groupKey: groupKeyForHolidays(ordered),
        seriesName: seriesDisplayName(ordered),
        holidays: ordered,
      };
    }),
  ];

  return groups.sort((a, b) =>
    a.holidays[0]!.date.localeCompare(b.holidays[0]!.date),
  );
}

function normalizeEntry(value: unknown): PakistanHolidayEmailEntry | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<PakistanHolidayEmailEntry>;
  const holidayDate = String(record.holidayDate ?? "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(holidayDate)) {
    return null;
  }
  const status: PakistanHolidayEmailStatus =
    record.status === "sent" ? "sent" : "scheduled";
  const holidayName = String(record.holidayName ?? "").trim();
  const holidayDates = Array.isArray(record.holidayDates)
    ? record.holidayDates
        .map((item) => String(item ?? "").trim())
        .filter((item) => /^\d{4}-\d{2}-\d{2}$/.test(item))
    : [];
  const holidayIds = Array.isArray(record.holidayIds)
    ? record.holidayIds.map((item) => String(item ?? "").trim()).filter(Boolean)
    : [];
  const dates = holidayDates.length > 0 ? holidayDates : [holidayDate];
  const seriesKey = pakistanHolidaySeriesKey(holidayName || "holiday");
  const groupKey =
    String(record.groupKey ?? "").trim() ||
    pakistanHolidayGroupKey(seriesKey, dates[0] ?? holidayDate);
  const audience: PakistanHolidayEmailAudience =
    record.audience === "client" ? "client" : "team";

  return {
    groupKey,
    audience,
    holidayDate: dates[0] ?? holidayDate,
    holidayDates: dates,
    holidayIds,
    holidayName,
    status,
    scheduledFor: String(record.scheduledFor ?? "").trim(),
    sentAt: String(record.sentAt ?? "").trim(),
  };
}

export function normalizePakistanHolidayEmailsStore(
  value: unknown,
): PakistanHolidayEmailsStore {
  if (!value || typeof value !== "object") {
    return createEmptyPakistanHolidayEmailsStore();
  }
  const record = value as Partial<PakistanHolidayEmailsStore>;
  const entries = Array.isArray(record.entries)
    ? record.entries
        .map((item) => normalizeEntry(item))
        .filter((item): item is PakistanHolidayEmailEntry => item !== null)
    : [];
  return { entries };
}

export function getPakistanHolidayEmailEntry(
  store: PakistanHolidayEmailsStore,
  groupKey: string,
  audience: PakistanHolidayEmailAudience = "team",
): PakistanHolidayEmailEntry | null {
  const key = groupKey.trim();
  return (
    store.entries.find(
      (item) => item.groupKey === key && item.audience === audience,
    ) ?? null
  );
}

export function getPakistanHolidayEmailEntryForHoliday(
  store: PakistanHolidayEmailsStore,
  holiday: Pick<HolidayEntry, "id" | "date" | "name">,
  audience: PakistanHolidayEmailAudience = "team",
): PakistanHolidayEmailEntry | null {
  const byId = store.entries.find(
    (item) =>
      item.audience === audience && item.holidayIds.includes(holiday.id),
  );
  if (byId) {
    return byId;
  }
  const seriesKey = pakistanHolidaySeriesKey(holiday.name);
  const slug = slugSeriesKey(seriesKey);
  const exactKey = pakistanHolidayGroupKey(seriesKey, holiday.date);
  const byExact = getPakistanHolidayEmailEntry(store, exactKey, audience);
  if (byExact) {
    return byExact;
  }
  // Mixed series (Arafa + Eid): groupKey is "2026:day-of-arafa+eid-ul-adha".
  return (
    store.entries.find((item) => {
      if (item.audience !== audience) {
        return false;
      }
      const dates =
        item.holidayDates.length > 0 ? item.holidayDates : [item.holidayDate];
      if (!dates.includes(holiday.date)) {
        return false;
      }
      const keyParts = item.groupKey.split(":")[1]?.split("+") ?? [];
      return keyParts.includes(slug);
    }) ?? null
  );
}

export function upsertPakistanHolidayEmailEntry(
  store: PakistanHolidayEmailsStore,
  entry: PakistanHolidayEmailEntry,
): PakistanHolidayEmailsStore {
  const without = store.entries.filter(
    (item) =>
      !(item.groupKey === entry.groupKey && item.audience === entry.audience),
  );
  return { entries: [...without, entry] };
}

function profileClientEmails(profile: {
  pocEmail?: string;
  documentRequestEmail?: string;
  financialsEmail?: string;
  ownerEmail?: string;
}): string[] {
  const emails: string[] = [];
  for (const raw of [
    profile.pocEmail,
    profile.documentRequestEmail,
    profile.financialsEmail,
    profile.ownerEmail,
  ]) {
    const email = normalizeEmail(String(raw ?? ""));
    if (email && isValidEmail(email)) {
      emails.push(email);
    }
  }
  return emails;
}

/** Unique client emails from Client Profiles (POC, document request, financials, owner). */
export function collectPakistanHolidayClientRecipientEmails(
  profiles: Array<{
    pocEmail?: string;
    documentRequestEmail?: string;
    financialsEmail?: string;
    ownerEmail?: string;
  }>,
): string[] {
  const emails = new Set<string>();
  for (const profile of profiles) {
    for (const email of profileClientEmails(profile)) {
      emails.add(email);
    }
  }
  return [...emails].sort((a, b) => a.localeCompare(b));
}

/**
 * Client emails grouped by profile time zone (9:00 AM local).
 * Missing/invalid zones fall back to Pakistan staff zone.
 * Each email is assigned to at most one group.
 */
export function collectPakistanHolidayClientRecipientGroups(
  profiles: Array<{
    pocEmail?: string;
    documentRequestEmail?: string;
    financialsEmail?: string;
    ownerEmail?: string;
    timeZone?: string;
  }>,
): PakistanHolidayClientRecipientGroup[] {
  const byZone = new Map<string, Set<string>>();
  const assigned = new Set<string>();

  for (const profile of profiles) {
    const emails = profileClientEmails(profile);
    if (emails.length === 0) {
      continue;
    }
    const timeZone =
      resolveClientTimeZone(profile.timeZone) ?? USER_TIME_ZONE;
    let bucket = byZone.get(timeZone);
    if (!bucket) {
      bucket = new Set<string>();
      byZone.set(timeZone, bucket);
    }
    for (const email of emails) {
      if (assigned.has(email)) {
        continue;
      }
      assigned.add(email);
      bucket.add(email);
    }
  }

  return [...byZone.entries()]
    .map(([timeZone, emails]) => ({
      timeZone,
      emails: [...emails].sort((a, b) => a.localeCompare(b)),
    }))
    .filter((group) => group.emails.length > 0)
    .sort((a, b) => a.timeZone.localeCompare(b.timeZone));
}

function shiftCalendarDate(date: string, deltaDays: number): string | null {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date.trim())) {
    return null;
  }
  const [year, month, day] = date.trim().split("-").map(Number);
  if (!year || !month || !day) {
    return null;
  }
  const utc = new Date(Date.UTC(year, month - 1, day));
  utc.setUTCDate(utc.getUTCDate() + deltaDays);
  const y = utc.getUTCFullYear();
  const m = String(utc.getUTCMonth() + 1).padStart(2, "0");
  const d = String(utc.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

/** Calendar day before YYYY-MM-DD (UTC date parts; holiday dates are date-only). */
export function dayBeforeHolidayDate(holidayDate: string): {
  year: number;
  monthIndex: number;
  day: number;
} | null {
  const before = shiftCalendarDate(holidayDate.trim(), -1);
  if (!before) {
    return null;
  }
  const [year, month, day] = before.split("-").map(Number);
  return {
    year: year!,
    monthIndex: month! - 1,
    day: day!,
  };
}

/** Sat/Sun for a date-only YYYY-MM-DD (UTC weekday). */
export function isWeekendHolidayDate(date: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date.trim())) {
    return false;
  }
  const [year, month, day] = date.trim().split("-").map(Number);
  const utc = new Date(Date.UTC(year!, month! - 1, day!));
  const weekday = utc.getUTCDay();
  return weekday === 0 || weekday === 6;
}

/**
 * Roll a calendar date back to Friday when it falls on Sat/Sun.
 */
export function adjustPakistanHolidayEmailSendDateToWeekday(
  date: string,
): string | null {
  let sendDate = date.trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(sendDate)) {
    return null;
  }
  while (isWeekendHolidayDate(sendDate)) {
    const previous = shiftCalendarDate(sendDate, -1);
    if (!previous) {
      return null;
    }
    sendDate = previous;
  }
  return sendDate;
}

/**
 * Team send calendar date: day before the holiday, rolled back to Friday when
 * that day falls on a weekend.
 */
export function pakistanHolidayEmailSendDate(holidayDate: string): string | null {
  const before = shiftCalendarDate(holidayDate.trim(), -1);
  if (!before) {
    return null;
  }
  return adjustPakistanHolidayEmailSendDateToWeekday(before);
}

/**
 * Client send calendar date: 7 business days before the holiday (Sat/Sun
 * skipped). The result is always a weekday.
 */
export function pakistanHolidayClientEmailSendDate(
  holidayDate: string,
): string | null {
  let current = holidayDate.trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(current)) {
    return null;
  }
  let remaining = PAKISTAN_HOLIDAY_CLIENT_EMAIL_LEAD_BUSINESS_DAYS;
  while (remaining > 0) {
    const previous = shiftCalendarDate(current, -1);
    if (!previous) {
      return null;
    }
    current = previous;
    if (!isWeekendHolidayDate(current)) {
      remaining -= 1;
    }
  }
  return current;
}

function sendAtFromYmd(sendDate: string, timeZone: string): Date | null {
  const [year, month, day] = sendDate.split("-").map(Number);
  if (!year || !month || !day) {
    return null;
  }
  return zonedWallTimeToUtc(
    year,
    month - 1,
    day,
    CLIENT_SEND_HOUR,
    0,
    timeZone,
  );
}

/** 9:00 AM Asia/Karachi on the last weekday before the holiday (team). */
export function pakistanHolidayEmailSendAt(holidayDate: string): Date | null {
  const sendDate = pakistanHolidayEmailSendDate(holidayDate);
  if (!sendDate) {
    return null;
  }
  return sendAtFromYmd(sendDate, USER_TIME_ZONE);
}

/**
 * 9:00 AM in the client's time zone, 7 business days before the holiday.
 */
export function pakistanHolidayClientEmailSendAt(
  holidayDate: string,
  timeZone: string = USER_TIME_ZONE,
): Date | null {
  const sendDate = pakistanHolidayClientEmailSendDate(holidayDate);
  if (!sendDate) {
    return null;
  }
  return sendAtFromYmd(sendDate, timeZone);
}

function holidayDateLabel(holidays: HolidayEntry[]): string {
  if (holidays.length === 0) {
    return "";
  }
  if (holidays.length === 1) {
    return formatHolidayDate(holidays[0]!.date);
  }
  const first = formatHolidayDate(holidays[0]!.date);
  const last = formatHolidayDate(holidays[holidays.length - 1]!.date);
  return first === last ? first : `${first} – ${last}`;
}

function holidayBulletBlock(holidays: HolidayEntry[]): string {
  return holidays
    .map((item) => `• ${item.name} — ${formatHolidayDate(item.date)}`)
    .join("\n");
}

export function fillPakistanHolidayEmailTemplate(
  subject: string,
  body: string,
  group: Pick<PakistanHolidayEmailGroup, "seriesName" | "holidays">,
): { subject: string; body: string } {
  const holidays = group.holidays;
  const nameLabel = group.seriesName;
  const dateLabel = holidayDateLabel(holidays);
  const year =
    Number(holidays[0]?.date.slice(0, 4)) || new Date().getFullYear();
  const periodDate = new Date(year, 0, 15);
  const bullets = holidayBulletBlock(holidays);

  const fillPlaceholders = (text: string) =>
    applyPeriodPlaceholder(
      text
        .replace(/\[Holiday Name\]/gi, nameLabel)
        .replace(/\[Holiday Date\]/gi, dateLabel),
      periodDate,
    );

  let nextBody = body.replace(
    /•\s*\[Holiday Name\]\s*[—–-]\s*\[Holiday Date\]/gi,
    bullets,
  );
  nextBody = fillPlaceholders(nextBody);
  const nextSubject = fillPlaceholders(subject);

  return {
    subject: nextSubject,
    body: nextBody,
  };
}

export type BuildPakistanHolidayEmailPlansOptions = {
  audience?: PakistanHolidayEmailAudience;
  /** Team: unused. Client: one Graph send per timezone group at 9:00 local. */
  recipientGroups?: PakistanHolidayClientRecipientGroup[];
  now?: Date;
};

/**
 * Build send plans: one email per holiday series (Eid Day 1–3 = one email).
 * Client audience expands to one plan per client time-zone group.
 */
export function buildPakistanHolidayEmailPlans(
  holidays: HolidayEntry[],
  store: PakistanHolidayEmailsStore,
  template: { subject: string; body: string },
  options: BuildPakistanHolidayEmailPlansOptions = {},
): PakistanHolidayEmailPlan[] {
  const audience = options.audience ?? "team";
  const now = options.now ?? new Date();
  const plans: PakistanHolidayEmailPlan[] = [];
  const soonestSchedule = now.getTime() + 60_000;
  const groups = groupPakistanHolidaysForEmail(holidays);
  const recipientGroups =
    audience === "client" ? (options.recipientGroups ?? []) : [];

  for (const group of groups) {
    const existing = getPakistanHolidayEmailEntry(
      store,
      group.groupKey,
      audience,
    );
    if (existing?.status === "sent" || existing?.status === "scheduled") {
      continue;
    }
    const firstDate = group.holidays[0]!.date;
    const lastDate = group.holidays[group.holidays.length - 1]!.date;
    // Skip if the whole series is already over (Pakistan date).
    const seriesEnd = new Date(`${lastDate}T23:59:59+05:00`);
    if (
      Number.isNaN(seriesEnd.getTime()) ||
      seriesEnd.getTime() < now.getTime()
    ) {
      continue;
    }

    const filled = fillPakistanHolidayEmailTemplate(
      template.subject,
      template.body,
      group,
    );
    const holidayDates = group.holidays.map((item) => item.date);
    const holidayIds = group.holidays.map((item) => item.id);

    const pushPlan = (
      sendAt: Date | null,
      extras?: { recipients?: string[]; timeZone?: string },
    ) => {
      const markAs: PakistanHolidayEmailStatus =
        sendAt && sendAt.getTime() >= soonestSchedule ? "scheduled" : "sent";
      plans.push({
        groupKey: group.groupKey,
        audience,
        holidayDate: firstDate,
        holidayDates,
        holidayIds,
        holidayName: group.seriesName,
        subject: filled.subject,
        body: filled.body,
        sendAt:
          markAs === "scheduled" && sendAt ? sendAt.toISOString() : null,
        markAs,
        recipients: extras?.recipients,
        timeZone: extras?.timeZone,
      });
    };

    if (audience === "client") {
      if (recipientGroups.length === 0) {
        continue;
      }
      for (const recipientGroup of recipientGroups) {
        const sendAt = pakistanHolidayClientEmailSendAt(
          firstDate,
          recipientGroup.timeZone,
        );
        if (!sendAt) {
          continue;
        }
        pushPlan(sendAt, {
          recipients: recipientGroup.emails,
          timeZone: recipientGroup.timeZone,
        });
      }
      continue;
    }

    const sendAt = pakistanHolidayEmailSendAt(firstDate);
    if (!sendAt) {
      continue;
    }
    pushPlan(sendAt);
  }

  return plans;
}

/** Promote scheduled rows whose send time has passed to Sent. */
export function promoteDueScheduledEntries(
  store: PakistanHolidayEmailsStore,
  now: Date = new Date(),
): { store: PakistanHolidayEmailsStore; changed: boolean } {
  let changed = false;
  const entries = store.entries.map((entry) => {
    if (entry.status !== "scheduled") {
      return entry;
    }
    const due = Date.parse(entry.scheduledFor);
    if (!due || due > now.getTime()) {
      return entry;
    }
    changed = true;
    return {
      ...entry,
      status: "sent" as const,
      sentAt: entry.sentAt || new Date(due).toISOString(),
    };
  });
  return { store: { entries }, changed };
}

export function displayPakistanHolidayEmailStatus(
  entry: PakistanHolidayEmailEntry | null,
): PakistanHolidayEmailStatusLabel {
  if (!entry) {
    return "—";
  }
  return entry.status === "sent" ? "Sent" : "Scheduled";
}

/** Status keyed by holiday id so multi-day series share one email status. */
export function buildPakistanHolidayEmailStatusByHolidayId(
  holidays: HolidayEntry[],
  store: PakistanHolidayEmailsStore,
  audience: PakistanHolidayEmailAudience = "team",
): Record<string, PakistanHolidayEmailStatusLabel> {
  const statusByHolidayId: Record<string, PakistanHolidayEmailStatusLabel> =
    {};
  for (const holiday of holidays) {
    statusByHolidayId[holiday.id] = displayPakistanHolidayEmailStatus(
      getPakistanHolidayEmailEntryForHoliday(store, holiday, audience),
    );
  }
  return statusByHolidayId;
}

export type PakistanHolidayEmailPreviewRow = {
  groupKey: string;
  holidayName: string;
  holidayDateLabel: string;
  teamSendDate: string | null;
  teamSendLabel: string;
  teamStatus: PakistanHolidayEmailStatusLabel;
  clientSendDate: string | null;
  clientSendLabel: string;
  clientStatus: PakistanHolidayEmailStatusLabel;
};

function formatPreviewSendLabel(
  sendDate: string | null,
  timeNote: string,
): string {
  if (!sendDate) {
    return "—";
  }
  return `${formatHolidayDate(sendDate)} · ${timeNote}`;
}

/**
 * Upcoming holiday email preview (team day-before + client 7 business days).
 * Status comes from the store once Graph has accepted the schedule.
 */
export function buildPakistanHolidayEmailPreview(
  holidays: HolidayEntry[],
  store: PakistanHolidayEmailsStore,
  now: Date = new Date(),
): PakistanHolidayEmailPreviewRow[] {
  const rows: PakistanHolidayEmailPreviewRow[] = [];
  const groups = groupPakistanHolidaysForEmail(holidays);

  for (const group of groups) {
    const firstDate = group.holidays[0]!.date;
    const lastDate = group.holidays[group.holidays.length - 1]!.date;
    const seriesEnd = new Date(`${lastDate}T23:59:59+05:00`);
    if (
      Number.isNaN(seriesEnd.getTime()) ||
      seriesEnd.getTime() < now.getTime()
    ) {
      continue;
    }

    const teamEntry = getPakistanHolidayEmailEntry(
      store,
      group.groupKey,
      "team",
    );
    const clientEntry = getPakistanHolidayEmailEntry(
      store,
      group.groupKey,
      "client",
    );
    const teamSendDate = pakistanHolidayEmailSendDate(firstDate);
    const clientSendDate = pakistanHolidayClientEmailSendDate(firstDate);

    rows.push({
      groupKey: group.groupKey,
      holidayName: group.seriesName,
      holidayDateLabel: holidayDateLabel(group.holidays),
      teamSendDate,
      teamSendLabel: formatPreviewSendLabel(
        teamSendDate,
        "9:00 AM Pakistan time",
      ),
      teamStatus: displayPakistanHolidayEmailStatus(teamEntry),
      clientSendDate,
      clientSendLabel: formatPreviewSendLabel(
        clientSendDate,
        "9:00 AM client time zone",
      ),
      clientStatus: displayPakistanHolidayEmailStatus(clientEntry),
    });
  }

  return rows;
}
