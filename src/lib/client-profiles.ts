import { normalizeYearEnd } from "@/lib/year-end";

export type ClientCountry = "US" | "CA" | "Not Sure";

export type PaymentType = "Cash" | "Accrual";

export const PAYMENT_TYPE_OPTIONS: PaymentType[] = ["Cash", "Accrual"];

export type YesNo = "Yes" | "No";

export const YES_NO_OPTIONS: YesNo[] = ["Yes", "No"];

export type PayrollReportsNeeded = YesNo;

export const PAYROLL_REPORTS_NEEDED_OPTIONS = YES_NO_OPTIONS;

export function normalizeYesNo(value: string): YesNo | "" {
  const lower = value.trim().toLowerCase();
  if (!lower) {
    return "";
  }
  if (
    lower === "yes" ||
    lower === "y" ||
    lower === "true" ||
    lower === "need reports" ||
    lower === "needs reports"
  ) {
    return "Yes";
  }
  if (
    lower === "no" ||
    lower === "n" ||
    lower === "false" ||
    lower === "do not need reports" ||
    lower === "doesn't need reports" ||
    lower === "none"
  ) {
    return "No";
  }
  return "";
}

export function normalizePayrollReportsNeeded(
  value: string,
): PayrollReportsNeeded | "" {
  return normalizeYesNo(value);
}

export const CLIENT_MEETING_FREQUENCY_OPTIONS = [
  "Monthly",
  "Quarterly",
  "None",
] as const;

export type ClientMeetingFrequency =
  (typeof CLIENT_MEETING_FREQUENCY_OPTIONS)[number] | "";

export function normalizeClientMeetingFrequency(
  value: string,
): ClientMeetingFrequency {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  const lower = trimmed.toLowerCase();
  if (lower === "monthly") {
    return "Monthly";
  }
  if (lower === "quarterly") {
    return "Quarterly";
  }
  if (lower === "none" || lower === "nil" || lower === "n/a") {
    return "None";
  }
  return (
    normalizeFixedOption(trimmed, CLIENT_MEETING_FREQUENCY_OPTIONS) || ""
  );
}

export function normalizePaymentType(value: string): PaymentType | "" {
  const lower = value.trim().toLowerCase();
  if (!lower) {
    return "";
  }
  if (
    lower === "cash" ||
    lower === "cash basis" ||
    lower === "cash-basis"
  ) {
    return "Cash";
  }
  if (
    lower === "accrual" ||
    lower === "accrual basis" ||
    lower === "accrual-basis"
  ) {
    return "Accrual";
  }
  return "";
}

/** US → USD ($), CA → CAD ($), Not Sure → None */
export function reportingCurrencyForCountry(country: ClientCountry): string {
  if (country === "CA") {
    return "CAD ($)";
  }
  if (country === "Not Sure") {
    return "None";
  }
  return "USD ($)";
}

export const SERVICE_FREQUENCY_OPTIONS = [
  "Monthly",
  "Quarterly",
  "Catch-up/Cleanup",
] as const;

export type ServiceFrequency = (typeof SERVICE_FREQUENCY_OPTIONS)[number] | "";

export const REPORTING_TIER_OPTIONS = [
  "Standard (financials by the 18th)",
  "Priority (by BD10 – BD15, for lender or board clients)",
  "Quarterly (within 20 – 30 days of quarter-end)",
] as const;

export type ReportingTier = (typeof REPORTING_TIER_OPTIONS)[number] | "";

function normalizeFixedOption<T extends string>(
  value: string,
  options: readonly T[],
): T | "" {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  const exact = options.find((option) => option === trimmed);
  if (exact) {
    return exact;
  }
  const lower = trimmed.toLowerCase();
  return options.find((option) => option.toLowerCase() === lower) ?? "";
}

export function normalizeServiceFrequency(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  const lower = trimmed.toLowerCase();
  if (lower === "monthly") {
    return "Monthly";
  }
  if (lower === "quarterly") {
    return "Quarterly";
  }
  if (
    lower === "catch-up/cleanup" ||
    lower === "catch-up" ||
    lower === "cleanup" ||
    lower === "catch up/cleanup" ||
    lower === "catchup/cleanup"
  ) {
    return "Catch-up/Cleanup";
  }
  return (
    normalizeFixedOption(trimmed, SERVICE_FREQUENCY_OPTIONS) || trimmed
  );
}

export function normalizeReportingTier(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  const lower = trimmed.toLowerCase();
  if (lower === "standard" || lower.startsWith("standard (")) {
    return "Standard (financials by the 18th)";
  }
  if (lower === "priority" || lower.startsWith("priority (")) {
    return "Priority (by BD10 – BD15, for lender or board clients)";
  }
  if (lower === "quarterly" || lower.startsWith("quarterly (")) {
    return "Quarterly (within 20 – 30 days of quarter-end)";
  }
  return normalizeFixedOption(trimmed, REPORTING_TIER_OPTIONS) || trimmed;
}

export const FINANCIAL_ACCOUNT_TYPE_OPTIONS = [
  "Bank",
  "Credit card",
  "Loan",
  "Merchant",
] as const;

export type FinancialAccountType =
  (typeof FINANCIAL_ACCOUNT_TYPE_OPTIONS)[number];

export const FINANCIAL_ACCOUNT_SOURCE_OPTIONS = [
  "Bank feed",
  "Needs statement",
] as const;

export type FinancialAccountSource =
  (typeof FINANCIAL_ACCOUNT_SOURCE_OPTIONS)[number];

export type FinancialAccount = {
  id: string;
  accountType: FinancialAccountType | "";
  institutionName: string;
  lastFour: string;
  source: FinancialAccountSource | "";
};

function makeFinancialAccountId(): string {
  return `acct-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function createEmptyFinancialAccount(): FinancialAccount {
  return {
    id: makeFinancialAccountId(),
    accountType: "",
    institutionName: "",
    lastFour: "",
    source: "",
  };
}

function normalizeFinancialAccountType(
  value: string,
): FinancialAccountType | "" {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  const lower = trimmed.toLowerCase();
  if (lower === "bank") {
    return "Bank";
  }
  if (
    lower === "credit card" ||
    lower === "creditcard" ||
    lower === "credit-card" ||
    lower === "cc"
  ) {
    return "Credit card";
  }
  if (lower === "loan") {
    return "Loan";
  }
  if (lower === "merchant" || lower === "merchant account") {
    return "Merchant";
  }
  return normalizeFixedOption(trimmed, FINANCIAL_ACCOUNT_TYPE_OPTIONS);
}

function normalizeFinancialAccountSource(
  value: string,
): FinancialAccountSource | "" {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  const lower = trimmed.toLowerCase();
  if (
    lower === "bank feed" ||
    lower === "bankfeed" ||
    lower === "feed" ||
    lower === "on feed"
  ) {
    return "Bank feed";
  }
  if (
    lower === "needs statement" ||
    lower === "needs statement from client" ||
    lower === "statement" ||
    lower === "client statement"
  ) {
    return "Needs statement";
  }
  return normalizeFixedOption(trimmed, FINANCIAL_ACCOUNT_SOURCE_OPTIONS);
}

function normalizeLastFour(value: string): string {
  return value.replace(/\D/g, "").slice(-4);
}

export function normalizeFinancialAccount(
  value: unknown,
): FinancialAccount | null {
  if (typeof value !== "object" || value === null) {
    return null;
  }
  const raw = value as Partial<FinancialAccount>;
  const institutionName = String(raw.institutionName ?? "").trim();
  const lastFour = normalizeLastFour(String(raw.lastFour ?? ""));
  const accountType = normalizeFinancialAccountType(
    String(raw.accountType ?? ""),
  );
  const source = normalizeFinancialAccountSource(String(raw.source ?? ""));
  const id = String(raw.id ?? "").trim() || makeFinancialAccountId();

  // Drop completely empty rows.
  if (!institutionName && !lastFour && !accountType && !source) {
    return null;
  }

  return {
    id,
    accountType,
    institutionName,
    lastFour,
    source,
  };
}

export function normalizeFinancialAccounts(value: unknown): FinancialAccount[] {
  if (!Array.isArray(value)) {
    if (typeof value === "string" && value.trim()) {
      return parseFinancialAccountsFromText(value);
    }
    return [];
  }
  return value
    .map((item) => normalizeFinancialAccount(item))
    .filter((item): item is FinancialAccount => item !== null);
}

/** One-line summary for tables/CSV, e.g. "Bank: Chase ···1234 (Bank feed)". */
export function formatFinancialAccountLabel(account: FinancialAccount): string {
  const type = account.accountType || "Account";
  const name = account.institutionName.trim() || "—";
  const lastFour = account.lastFour.trim()
    ? `···${account.lastFour.trim()}`
    : "···????";
  const source = account.source || "source unset";
  return `${type}: ${name} ${lastFour} (${source})`;
}

export function formatFinancialAccountsList(
  accounts: FinancialAccount[],
): string {
  return accounts.map(formatFinancialAccountLabel).join("\n");
}

/**
 * Parse CSV/text back into accounts.
 * Accepts one account per line, or semicolon-separated (legacy):
 * "Bank: Chase ···1234 (Bank feed)\nCredit card: Amex ···5678 (Needs statement)"
 */
export function parseFinancialAccountsFromText(
  value: string,
): FinancialAccount[] {
  const trimmed = value.trim();
  if (!trimmed || trimmed === "—") {
    return [];
  }
  const parts = trimmed
    .split(/\s*(?:;|\r?\n)\s*/)
    .map((part) => part.trim())
    .filter(Boolean);
  const accounts: FinancialAccount[] = [];
  for (const part of parts) {
    const match = part.match(
      /^(.+?):\s*(.+?)\s*(?:···|…|\.\.\.)(\d{1,4})\s*\((.+)\)$/i,
    );
    if (match) {
      const account = normalizeFinancialAccount({
        accountType: match[1],
        institutionName: match[2],
        lastFour: match[3],
        source: match[4],
      });
      if (account) {
        accounts.push(account);
      }
      continue;
    }
    const loose = normalizeFinancialAccount({
      institutionName: part,
    });
    if (loose) {
      accounts.push(loose);
    }
  }
  return accounts;
}

export type Bd1TemplateId = "template-1" | "template-2";

export const BD1_TEMPLATE_OPTIONS: {
  id: Bd1TemplateId;
  label: string;
}[] = [
  { id: "template-1", label: "Template 1" },
  { id: "template-2", label: "Template 2" },
];

export function normalizeBd1Template(value: string): Bd1TemplateId {
  return value.trim().toLowerCase() === "template-2"
    ? "template-2"
    : "template-1";
}

/** Maps Client Profile BD1 Template to the BD1 request-type menu id. */
export function bd1TemplateToRequestTypeId(
  template: Bd1TemplateId,
): "bank-credit-card" | "other-statements" {
  return template === "template-2" ? "other-statements" : "bank-credit-card";
}

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
  reportingCurrency: string;
  /** Which BD1 request package this client uses. Default Template 1. */
  bd1Template: Bd1TemplateId;
  serviceFrequency: string;
  reportingTier: string;
  paymentType: string;
  financialAccounts: FinancialAccount[];
  hasUsdAndCadAccounts: string;
  meetingFrequency: string;
  onCatchUpPlan: string;
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
  payrollProvider: string;
  payrollReportsNeeded: string;
  payrollAccountNumbers: string;
  ownerName: string;
  ownerEmail: string;
  ownerPhone: string;
  pointOfContact: string;
  pocEmail: string;
  pocPhone: string;
  documentRequestContact: string;
  documentRequestEmail: string;
  documentRequestPhone: string;
  financialsContact: string;
  financialsEmail: string;
  financialsPhone: string;
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
    reportingCurrency: reportingCurrencyForCountry(country),
    bd1Template: "template-1",
    serviceFrequency: "",
    reportingTier: "",
    paymentType: "",
    financialAccounts: [],
    hasUsdAndCadAccounts: "",
    meetingFrequency: "",
    onCatchUpPlan: "",
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
    payrollProvider: "",
    payrollReportsNeeded: "",
    payrollAccountNumbers: "",
    ownerName: "",
    ownerEmail: "",
    ownerPhone: "",
    pointOfContact: "",
    pocEmail: "",
    pocPhone: "",
    documentRequestContact: "",
    documentRequestEmail: "",
    documentRequestPhone: "",
    financialsContact: "",
    financialsEmail: "",
    financialsPhone: "",
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
    reportingCurrency: reportingCurrencyForCountry(country),
    bd1Template: normalizeBd1Template(String(value.bd1Template ?? "")),
    serviceFrequency: normalizeServiceFrequency(
      String(value.serviceFrequency ?? ""),
    ),
    reportingTier: normalizeReportingTier(String(value.reportingTier ?? "")),
    paymentType: normalizePaymentType(String(value.paymentType ?? "")),
    financialAccounts: normalizeFinancialAccounts(value.financialAccounts),
    hasUsdAndCadAccounts: normalizeYesNo(
      String(value.hasUsdAndCadAccounts ?? ""),
    ),
    meetingFrequency: normalizeClientMeetingFrequency(
      String(value.meetingFrequency ?? ""),
    ),
    onCatchUpPlan: normalizeYesNo(String(value.onCatchUpPlan ?? "")),
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
    payrollProvider: String(value.payrollProvider ?? "").trim(),
    payrollReportsNeeded: normalizePayrollReportsNeeded(
      String(value.payrollReportsNeeded ?? ""),
    ),
    payrollAccountNumbers: String(value.payrollAccountNumbers ?? "").trim(),
    ownerName: String(value.ownerName ?? "").trim(),
    ownerEmail: String(value.ownerEmail ?? "").trim(),
    ownerPhone: String(value.ownerPhone ?? "").trim(),
    pointOfContact: String(value.pointOfContact ?? "").trim(),
    pocEmail: String(value.pocEmail ?? "").trim(),
    pocPhone: String(value.pocPhone ?? "").trim(),
    documentRequestContact: String(value.documentRequestContact ?? "").trim(),
    documentRequestEmail: String(value.documentRequestEmail ?? "").trim(),
    documentRequestPhone: String(value.documentRequestPhone ?? "").trim(),
    financialsContact: String(value.financialsContact ?? "").trim(),
    financialsEmail: String(value.financialsEmail ?? "").trim(),
    financialsPhone: String(value.financialsPhone ?? "").trim(),
    createdAt:
      typeof value.createdAt === "string" && value.createdAt
        ? value.createdAt
        : new Date().toISOString(),
  };
}
