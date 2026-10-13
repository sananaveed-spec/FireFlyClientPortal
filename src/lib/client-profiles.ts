import { normalizeYearEnd } from "@/lib/year-end";

export type ClientCountry = "US" | "CA" | "Not Sure";

export type ClientProfileEntry = {
  id: string;
  uniqueId: string;
  legalName: string;
  operatingName: string;
  website: string;
  portalLink: string;
  industry: string;
  businessActivity: string;
  entityType: string;
  fiscalYearEnd: string;
  country: ClientCountry;
  state: string;
  streetAddress: string;
  businessAccountNumber: string;
  timeZone: string;
  ssn: string;
  doi: string;
  ein: string;
  ctTaxRegistrationNo: string;
  canadianBusinessNumber: string;
  gstHst: string;
  provincialSalesTaxNumbers: string;
  payrollAccountNumbers: string;
  ownerName: string;
  ownerEmail: string;
  ownerPhone: string;
  pointOfContact: string;
  pocEmail: string;
  pocPhone: string;
  createdAt: string;
};

export type ClientProfileInput = Omit<ClientProfileEntry, "id" | "createdAt">;

export function countryCode(country: ClientCountry): "US" | "CA" | "NS" {
  if (country === "CA") {
    return "CA";
  }
  if (country === "Not Sure") {
    return "NS";
  }
  return "US";
}

export function formatUniqueIdDate(date = new Date()): string {
  // UTC keeps the date stable regardless of server timezone.
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}${month}${day}`;
}

/**
 * Format: {CC}-{YYYYMMDD}-{NNN}
 * Examples: US-20261006-001, CA-20261006-002, NS-20261006-001
 *
 * Sequence NNN continues across dates for the same country code
 * (e.g. after US-20261006-010, next is US-20261008-011).
 */
export function generateUniqueId(
  existing: string[] = [],
  country: ClientCountry = "US",
  date = new Date(),
): string {
  const cc = countryCode(country);
  const prefix = `${cc}-${formatUniqueIdDate(date)}-`;
  const countryPrefix = `${cc}-`;
  const used = new Set(
    existing.map((value) => value.trim().toLowerCase()).filter(Boolean),
  );

  let maxSeq = 0;
  for (const id of existing) {
    const trimmed = id.trim();
    if (!trimmed.toUpperCase().startsWith(countryPrefix)) {
      continue;
    }
    // Prefer strict {CC}-{YYYYMMDD}-{NNN}; otherwise use trailing digits.
    const match = trimmed.match(/^(?:US|CA|NS)-\d{8}-(\d{3})$/i);
    const seq = Number.parseInt(
      match?.[1] ?? trimmed.slice(trimmed.lastIndexOf("-") + 1),
      10,
    );
    if (Number.isFinite(seq) && seq > maxSeq) {
      maxSeq = seq;
    }
  }

  for (let seq = maxSeq + 1; seq < maxSeq + 10000; seq += 1) {
    const candidate = `${prefix}${String(seq).padStart(3, "0")}`;
    if (!used.has(candidate.toLowerCase())) {
      return candidate;
    }
  }

  return `${prefix}${String(maxSeq + 1).padStart(3, "0")}-${Math.random()
    .toString(36)
    .slice(2, 4)
    .toUpperCase()}`;
}

export function matchesUniqueIdFormat(uniqueId: string): boolean {
  return /^(US|CA|NS)-\d{8}-\d{3}$/i.test(uniqueId.trim());
}

/** Rebuild Unique IDs for profiles that are not in {CC}-{YYYYMMDD}-{NNN} form. */
export function migrateUniqueIds(
  profiles: ClientProfileEntry[],
): { profiles: ClientProfileEntry[]; changed: boolean } {
  if (profiles.length === 0) {
    return { profiles, changed: false };
  }

  const needsMigration = profiles.some(
    (profile) => !matchesUniqueIdFormat(profile.uniqueId),
  );
  if (!needsMigration) {
    return { profiles, changed: false };
  }

  const next = profiles.map((profile) => ({ ...profile }));
  const assigned = next
    .filter((profile) => matchesUniqueIdFormat(profile.uniqueId))
    .map((profile) => profile.uniqueId);

  const order = next
    .map((profile, index) => ({ profile, index }))
    .sort((a, b) => {
      const timeA = new Date(a.profile.createdAt).getTime();
      const timeB = new Date(b.profile.createdAt).getTime();
      if (timeA !== timeB) {
        return timeA - timeB;
      }
      return a.profile.id.localeCompare(b.profile.id);
    });

  let changed = false;
  for (const { profile, index } of order) {
    if (matchesUniqueIdFormat(next[index].uniqueId)) {
      continue;
    }
    const created = new Date(profile.createdAt);
    const uniqueId = generateUniqueId(
      assigned,
      profile.country,
      Number.isNaN(created.getTime()) ? new Date() : created,
    );
    next[index] = { ...next[index], uniqueId };
    assigned.push(uniqueId);
    changed = true;
  }

  return { profiles: next, changed };
}

export function createEmptyClientProfileForm(
  existingUniqueIds: string[] = [],
  country: ClientCountry = "US",
): ClientProfileInput {
  return {
    uniqueId: generateUniqueId(existingUniqueIds, country),
    legalName: "",
    operatingName: "",
    website: "",
    portalLink: "",
    industry: "",
    businessActivity: "",
    entityType: "",
    fiscalYearEnd: "",
    country,
    state: "",
    streetAddress: "",
    businessAccountNumber: "",
    timeZone: "",
    ssn: "",
    doi: "",
    ein: "",
    ctTaxRegistrationNo: "",
    canadianBusinessNumber: "",
    gstHst: "",
    provincialSalesTaxNumbers: "",
    payrollAccountNumbers: "",
    ownerName: "",
    ownerEmail: "",
    ownerPhone: "",
    pointOfContact: "",
    pocEmail: "",
    pocPhone: "",
  };
}

export function normalizeClientProfile(
  value: Partial<ClientProfileEntry> & { id: string },
): ClientProfileEntry {
  const rawCountry = String(value.country ?? "").trim().toLowerCase();
  const country: ClientCountry =
    rawCountry === "ca" ||
    rawCountry === "canada" ||
    rawCountry === "canadian"
      ? "CA"
      : rawCountry === "not sure" || rawCountry === "ns"
        ? "Not Sure"
        : "US";

  return {
    id: value.id,
    uniqueId: String(value.uniqueId ?? "").trim(),
    legalName: String(value.legalName ?? "").trim(),
    operatingName: String(value.operatingName ?? "").trim(),
    website: String(value.website ?? "").trim(),
    portalLink: String(value.portalLink ?? "").trim(),
    industry: String(value.industry ?? "").trim(),
    businessActivity: String(value.businessActivity ?? "").trim(),
    entityType: String(value.entityType ?? "").trim(),
    fiscalYearEnd: normalizeYearEnd(String(value.fiscalYearEnd ?? "")),
    country,
    state: String(value.state ?? "").trim(),
    streetAddress: String(value.streetAddress ?? "").trim(),
    businessAccountNumber: String(value.businessAccountNumber ?? "").trim(),
    timeZone: String(value.timeZone ?? "").trim(),
    ssn: String(value.ssn ?? "").trim(),
    doi: String(value.doi ?? "").trim(),
    ein: String(value.ein ?? "").trim(),
    ctTaxRegistrationNo: String(value.ctTaxRegistrationNo ?? "").trim(),
    canadianBusinessNumber: String(value.canadianBusinessNumber ?? "").trim(),
    gstHst: String(value.gstHst ?? "").trim(),
    provincialSalesTaxNumbers: String(
      value.provincialSalesTaxNumbers ?? "",
    ).trim(),
    payrollAccountNumbers: String(value.payrollAccountNumbers ?? "").trim(),
    ownerName: String(value.ownerName ?? "").trim(),
    ownerEmail: String(value.ownerEmail ?? "").trim(),
    ownerPhone: String(value.ownerPhone ?? "").trim(),
    pointOfContact: String(value.pointOfContact ?? "").trim(),
    pocEmail: String(value.pocEmail ?? "").trim(),
    pocPhone: String(value.pocPhone ?? "").trim(),
    createdAt:
      typeof value.createdAt === "string" && value.createdAt
        ? value.createdAt
        : new Date().toISOString(),
  };
}
