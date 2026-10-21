/** Staff / “your” timezone for email scheduling (Pakistan). */
export const USER_TIME_ZONE = "Asia/Karachi";

/** Default send hour in the client’s local time. */
export const CLIENT_SEND_HOUR = 9;

const TIME_ZONE_ALIASES: Record<string, string> = {
  est: "America/New_York",
  edt: "America/New_York",
  eastern: "America/New_York",
  "us/eastern": "America/New_York",
  cst: "America/Chicago",
  cdt: "America/Chicago",
  central: "America/Chicago",
  "us/central": "America/Chicago",
  mst: "America/Denver",
  mdt: "America/Denver",
  mountain: "America/Denver",
  "us/mountain": "America/Denver",
  pst: "America/Los_Angeles",
  pdt: "America/Los_Angeles",
  pacific: "America/Los_Angeles",
  "us/pacific": "America/Los_Angeles",
  akst: "America/Anchorage",
  hst: "Pacific/Honolulu",
  utc: "UTC",
  gmt: "UTC",
  pkt: "Asia/Karachi",
  pakistan: "Asia/Karachi",
  "asia/karachi": "Asia/Karachi",
  vietnam: "Asia/Ho_Chi_Minh",
  "vietnam 2 hours ahead": "Asia/Ho_Chi_Minh",
  "ho chi minh": "Asia/Ho_Chi_Minh",
  ict: "Asia/Ho_Chi_Minh",
};

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

/** Map free-form client profile labels (EST, PST, …) to an IANA zone. */
export function resolveClientTimeZone(label: string | undefined | null): string | null {
  const raw = String(label ?? "").trim();
  if (!raw) {
    return null;
  }

  const lower = raw.toLowerCase();
  if (TIME_ZONE_ALIASES[lower]) {
    return TIME_ZONE_ALIASES[lower];
  }

  // Allow pasting a real IANA id.
  if (/^[A-Za-z]+\/[A-Za-z0-9_+\-]+$/.test(raw)) {
    try {
      Intl.DateTimeFormat(undefined, { timeZone: raw });
      return raw;
    } catch {
      // fall through
    }
  }

  // Partial contains: "EST (New York)"
  for (const [alias, iana] of Object.entries(TIME_ZONE_ALIASES)) {
    if (lower.includes(alias)) {
      return iana;
    }
  }

  return null;
}

type ZoneParts = {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
};

function getZonedParts(date: Date, timeZone: string): ZoneParts {
  const dtf = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  });
  const map: Record<string, string> = {};
  for (const part of dtf.formatToParts(date)) {
    if (part.type !== "literal") {
      map[part.type] = part.value;
    }
  }
  return {
    year: Number(map.year),
    month: Number(map.month),
    day: Number(map.day),
    hour: Number(map.hour === "24" ? "0" : map.hour),
    minute: Number(map.minute),
    second: Number(map.second),
  };
}

/** Instant for a wall-clock date/time in `timeZone`. */
export function zonedWallTimeToUtc(
  year: number,
  monthIndex: number,
  day: number,
  hour: number,
  minute: number,
  timeZone: string,
): Date {
  // Iterate: treat desired wall as UTC, then correct by zone offset.
  let utcMs = Date.UTC(year, monthIndex, day, hour, minute, 0);
  for (let i = 0; i < 3; i += 1) {
    const parts = getZonedParts(new Date(utcMs), timeZone);
    const asIfUtc = Date.UTC(
      parts.year,
      parts.month - 1,
      parts.day,
      parts.hour,
      parts.minute,
      parts.second,
    );
    const desired = Date.UTC(year, monthIndex, day, hour, minute, 0);
    utcMs += desired - asIfUtc;
  }
  return new Date(utcMs);
}

export function formatDateTimeLocalInTimeZone(
  date: Date,
  timeZone: string,
): string {
  const p = getZonedParts(date, timeZone);
  return `${p.year}-${pad2(p.month)}-${pad2(p.day)}T${pad2(p.hour)}:${pad2(p.minute)}`;
}

/** Parse `YYYY-MM-DDTHH:mm` as wall time in `timeZone` → UTC Date. */
export function parseDateTimeLocalInTimeZone(
  value: string,
  timeZone: string,
): Date | null {
  const trimmed = value.trim();
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(trimmed)) {
    return null;
  }
  const [datePart, timePart] = trimmed.split("T");
  const [year, month, day] = datePart.split("-").map(Number);
  const [hour, minute] = timePart.split(":").map(Number);
  const date = zonedWallTimeToUtc(
    year,
    month - 1,
    day,
    hour,
    minute,
    timeZone,
  );
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return date;
}

export function formatClientSendAtLabel(
  date: Date,
  clientTimeZoneIana: string,
  clientTimeZoneLabel: string,
): string {
  const formatted = new Intl.DateTimeFormat("en-GB", {
    timeZone: clientTimeZoneIana,
    weekday: "long",
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  }).format(date);
  const label = clientTimeZoneLabel.trim() || clientTimeZoneIana;
  return `${formatted} (${label})`;
}

export function formatUserSendAtLabel(date: Date): string {
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: USER_TIME_ZONE,
    weekday: "long",
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  }).format(date);
}

export type ClientNineAmSchedule = {
  /** UTC instant for 9:00 client time (future). */
  instant: Date;
  clientTimeZoneIana: string;
  clientLabel: string;
  /** datetime-local value in Pakistan (your) timezone — use for scheduling. */
  userDateTimeLocal: string;
  userLabel: string;
};

export const CLIENT_TIME_ZONE_REQUIRED_MESSAGE =
  "Set the client time zone in Client Profile.";

/**
 * 9:00 AM on the schedule BD date in the client timezone, converted to
 * Pakistan for the editable “Scheduled send (Pakistan)” field.
 * Always uses that BD calendar date (does not skip past days).
 * Returns null when the client time zone is missing or unrecognized.
 */
export function getClientNineAmSchedule(
  scheduleDate: Date | null,
  clientTimeZoneLabel: string | undefined | null,
): ClientNineAmSchedule | null {
  if (!scheduleDate) {
    return null;
  }
  const clientIana = resolveClientTimeZone(clientTimeZoneLabel);
  if (!clientIana) {
    return null;
  }
  const year = scheduleDate.getFullYear();
  const monthIndex = scheduleDate.getMonth();
  const day = scheduleDate.getDate();
  const instant = zonedWallTimeToUtc(
    year,
    monthIndex,
    day,
    CLIENT_SEND_HOUR,
    0,
    clientIana,
  );
  return {
    instant,
    clientTimeZoneIana: clientIana,
    clientLabel: formatClientSendAtLabel(
      instant,
      clientIana,
      clientTimeZoneLabel?.trim() || clientIana,
    ),
    userDateTimeLocal: formatDateTimeLocalInTimeZone(instant, USER_TIME_ZONE),
    userLabel: `${formatUserSendAtLabel(instant)} (Pakistan)`,
  };
}
