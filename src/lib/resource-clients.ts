import { resolveYearlyComplianceTimelines } from "@/lib/yearly-compliance";

export type BookkeepingCategory = "ACC1" | "ACC2" | "ACC3";

/** US → ACC1, CA → ACC2, Not Sure → ACC3 */
export function bookkeepingCategoryForCountry(
  country: string | undefined | null,
): BookkeepingCategory {
  const normalized = String(country ?? "").trim().toLowerCase();
  if (
    normalized === "ca" ||
    normalized === "canada" ||
    normalized === "canadian"
  ) {
    return "ACC2";
  }
  if (normalized === "not sure" || normalized === "ns") {
    return "ACC3";
  }
  return "ACC1";
}

export type ResourceClientEntry = {
  id: string;
  bookkeepingCategory: BookkeepingCategory;
  clientName: string;
  industry: string;
  status: string;
  supervisor: string;
  bookkeeperAccountant: string[];
  bookkeeping: string[];
  payroll: string;
  payrollAndSalesTaxes: string[];
  deliverables: string[];
  yearlyCompliance: string[];
  meetingRequirements: string;
  meetingFrequency: string;
  bookkeepingDeadline: string;
  payrollDeadline: string;
  payrollTaxesDeadlines: string;
  deliverablesTimelines: string;
  yearlyComplianceTimelines: string;
  createdAt: string;
};

export type ResourceClientInput = Omit<ResourceClientEntry, "id" | "createdAt">;

function asStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map((item) => String(item ?? "").trim())
    .filter(Boolean);
}

export function formatMultiValue(values: string[]): string {
  return values.filter(Boolean).join(" + ");
}

export function createEmptyResourceClientForm(
  defaultStatus = "",
): ResourceClientInput {
  return {
    bookkeepingCategory: "ACC1",
    clientName: "",
    industry: "",
    status: defaultStatus,
    supervisor: "",
    bookkeeperAccountant: [],
    bookkeeping: [],
    payroll: "",
    payrollAndSalesTaxes: [],
    deliverables: [],
    yearlyCompliance: [],
    meetingRequirements: "",
    meetingFrequency: "",
    bookkeepingDeadline: "",
    payrollDeadline: "",
    payrollTaxesDeadlines: "",
    deliverablesTimelines: "",
    yearlyComplianceTimelines: "",
  };
}

export function normalizeResourceClient(
  value: Partial<ResourceClientEntry> & { id: string },
): ResourceClientEntry {
  const bookkeepingCategory =
    value.bookkeepingCategory === "ACC1" ||
    value.bookkeepingCategory === "ACC2" ||
    value.bookkeepingCategory === "ACC3"
      ? value.bookkeepingCategory
      : "ACC1";

  const yearlyCompliance = asStringArray(value.yearlyCompliance);
  const yearlyComplianceTimelines =
    resolveYearlyComplianceTimelines(yearlyCompliance) ||
    String(value.yearlyComplianceTimelines ?? "").trim();

  return {
    id: value.id,
    bookkeepingCategory,
    clientName: String(value.clientName ?? "").trim(),
    industry: String(value.industry ?? "").trim(),
    status: String(value.status ?? "").trim(),
    supervisor: String(value.supervisor ?? "").trim(),
    bookkeeperAccountant: asStringArray(value.bookkeeperAccountant),
    bookkeeping: asStringArray(value.bookkeeping),
    payroll: String(value.payroll ?? "").trim(),
    payrollAndSalesTaxes: asStringArray(value.payrollAndSalesTaxes),
    deliverables: asStringArray(value.deliverables),
    yearlyCompliance,
    meetingRequirements: String(value.meetingRequirements ?? "").trim(),
    meetingFrequency: String(value.meetingFrequency ?? "").trim(),
    bookkeepingDeadline: String(value.bookkeepingDeadline ?? "").trim(),
    payrollDeadline: String(value.payrollDeadline ?? "").trim(),
    payrollTaxesDeadlines: String(value.payrollTaxesDeadlines ?? "").trim(),
    deliverablesTimelines: String(value.deliverablesTimelines ?? "").trim(),
    yearlyComplianceTimelines,
    createdAt:
      typeof value.createdAt === "string" && value.createdAt
        ? value.createdAt
        : new Date().toISOString(),
  };
}

export function resolveMeetingFrequency(
  meetingRequirements: string,
  currentFrequency: string,
): { frequency: string; locked: boolean } {
  const normalized = meetingRequirements.trim().toLowerCase();
  if (normalized === "nil") {
    return { frequency: "NIL", locked: true };
  }
  if (normalized === "on demand") {
    return { frequency: "On Demand", locked: true };
  }
  if (normalized === "yes") {
    return { frequency: currentFrequency, locked: false };
  }
  return { frequency: "", locked: true };
}
