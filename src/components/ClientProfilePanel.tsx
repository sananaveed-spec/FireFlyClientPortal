"use client";

import { useMsal } from "@azure/msal-react";
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type FormEvent,
  type ReactNode,
} from "react";
import { createPortal } from "react-dom";
import { useAccess } from "@/auth/accessContext";
import {
  IgnitionProposalViewer,
  type IgnitionProposalViewerTarget,
} from "@/components/IgnitionProposalViewer";
import {
  BD1_TEMPLATE_OPTIONS,
  createEmptyClientProfileForm,
  createEmptyFinancialAccount,
  FINANCIAL_ACCOUNT_SOURCE_OPTIONS,
  FINANCIAL_ACCOUNT_TYPE_OPTIONS,
  formatFinancialAccountsList,
  generateUniqueId,
  normalizeBd1Template,
  normalizeClientProfile,
  normalizeFinancialAccounts,
  CLIENT_MEETING_FREQUENCY_OPTIONS,
  normalizeClientMeetingFrequency,
  normalizePaymentType,
  normalizePayrollReportsNeeded,
  normalizeReportingTier,
  normalizeServiceFrequency,
  normalizeYesNo,
  parseFinancialAccountsFromText,
  PAYMENT_TYPE_OPTIONS,
  PAYROLL_REPORTS_NEEDED_OPTIONS,
  reportingCurrencyForCountry,
  YES_NO_OPTIONS,
  type Bd1TemplateId,
  type ClientCountry,
  type ClientMeetingFrequency,
  type ClientProfileEntry,
  type ClientProfileInput,
  type FinancialAccount,
  type FinancialAccountSource,
  type FinancialAccountType,
  type PaymentType,
  type PayrollReportsNeeded,
  type YesNo,
} from "@/lib/client-profiles";
import { portalFetch } from "@/lib/portal-api";
import {
  DataListToolbar,
  downloadCsv,
  parseCsv,
  uniqueSortedValues,
} from "@/components/DataListToolbar";
import { RowActionsMenu } from "@/components/RowActionsMenu";
import { SearchableSelect } from "@/components/SearchableSelect";
import {
  buildChangeDetails,
  snapshotFromLabeledFields,
} from "@/lib/history";
import {
  getActorFromAccount,
  logPortalAction,
} from "@/lib/history-client";
import {
  createLookupItem,
  getLookupOptions,
  type LookupDataStore,
} from "@/lib/lookup-data";
import { getRegionsForCountry } from "@/lib/regions";
import {
  formatYearEnd,
  normalizeYearEnd,
  parseYearEnd,
  YEAR_END_DAYS,
  YEAR_END_MONTHS,
} from "@/lib/year-end";
import {
  clientSharePointSiteDisplayName,
  createClientSharePointSite,
  isClientSharePointSite,
  listSharePointSitesForPicker,
  parseMemberEmails,
  type SharePointSiteOption,
} from "@/lib/sharepoint-graph";
import type { SearchableOption } from "@/components/SearchableSelect";

function siteMatchesLegalName(siteName: string, legalName: string): boolean {
  const site = siteName.trim().toLowerCase();
  const legal = legalName.trim().toLowerCase();
  if (!site || !legal) {
    return false;
  }
  if (site.includes(legal)) {
    return true;
  }
  const stripped = site.replace(/^clients?\s*-\s*/i, "").trim();
  return stripped.includes(legal) || legal.includes(stripped);
}

async function persistLookupData(data: LookupDataStore) {
  const response = await portalFetch("/api/lookup-data", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  const payload = (await response.json()) as LookupDataStore & {
    error?: string;
  };
  if (!response.ok) {
    throw new Error(payload.error || "Failed to save lookup data.");
  }
  return payload;
}

const PROFILE_SORT_OPTIONS = [
  { id: "nameAsc", label: "Alphabetically — Legal Name" },
  { id: "idAsc", label: "Alphabetically — Unique ID" },
  { id: "newest", label: "Newest first" },
  { id: "oldest", label: "Oldest first" },
];

const MIN_PAGE_SIZE = 1;
const FALLBACK_HEADER_HEIGHT = 42;
const FALLBACK_ROW_HEIGHT = 40;

const TABLE_COLUMNS: { key: keyof ClientProfileEntry; label: string }[] = [
  { key: "uniqueId", label: "Unique ID" },
  { key: "legalName", label: "Legal Name" },
  { key: "operatingName", label: "Operating Name" },
  { key: "industry", label: "Industry" },
  { key: "businessActivity", label: "Business Activity" },
  { key: "entityType", label: "Entity Type" },
  { key: "serviceFrequency", label: "Service Frequency" },
  { key: "reportingTier", label: "Reporting Tier" },
  { key: "paymentType", label: "Payment Type" },
  { key: "financialAccounts", label: "Financial Accounts" },
  { key: "hasUsdAndCadAccounts", label: "USD and CAD Accounts" },
  { key: "meetingFrequency", label: "Meeting Frequency" },
  { key: "onCatchUpPlan", label: "On Catch-up Plan" },
  { key: "country", label: "Country" },
  { key: "reportingCurrency", label: "Reporting Currency" },
  { key: "state", label: "State/Province" },
  { key: "streetAddress", label: "Business Address" },
  { key: "businessAccountNumber", label: "Business Account Number" },
  { key: "timeZone", label: "Time Zone" },
  { key: "ssn", label: "SSN" },
  { key: "doi", label: "DOI (MM/DD/YYYY)" },
  { key: "ein", label: "FEIN" },
  { key: "ctTaxRegistrationNo", label: "Provincial Sales Tax Numbers" },
  { key: "fiscalYearEnd", label: "Year End" },
  { key: "website", label: "Website" },
  { key: "portalLink", label: "Portal Link" },
  { key: "activeProposalReference", label: "Active Proposal" },
  { key: "canadianBusinessNumber", label: "Canadian Business Number" },
  { key: "gstHst", label: "GST/HST" },
  { key: "payrollProvider", label: "Payroll Provider" },
  { key: "payrollReportsNeeded", label: "Payroll Reports Needed" },
  { key: "payrollAccountNumbers", label: "Payroll Account Numbers" },
  { key: "ownerName", label: "Owner Name" },
  { key: "ownerEmail", label: "Owner Email" },
  { key: "ownerPhone", label: "Owner Phone Number" },
  { key: "pointOfContact", label: "Point Of Contact" },
  { key: "pocEmail", label: "POC Email" },
  { key: "pocPhone", label: "POC Phone Number" },
  { key: "documentRequestContact", label: "Document Request Contact" },
  { key: "documentRequestEmail", label: "Document Request Email" },
  { key: "documentRequestPhone", label: "Document Request Phone" },
  { key: "financialsContact", label: "Financials Contact" },
  { key: "financialsEmail", label: "Financials Email" },
  { key: "financialsPhone", label: "Financials Phone" },
];

/** Display YYYY-MM-DD as MM/DD/YYYY (e.g. 06/07/1977). */
function formatProfileDate(value: string): string {
  const match = value.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return value.trim();
  }
  return `${match[2]}/${match[3]}/${match[1]}`;
}

/** Accept YYYY-MM-DD or MM/DD/YYYY for DOI storage. */
function parseImportDate(value: string): string {
  const trimmed = value.trim();
  if (!trimmed || trimmed === "—" || trimmed === "***") {
    return "";
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    return trimmed;
  }
  const match = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) {
    return trimmed;
  }
  return `${match[3]}-${match[1].padStart(2, "0")}-${match[2].padStart(2, "0")}`;
}

function parseImportCountry(value: string): ClientCountry {
  const normalized = value.trim().toLowerCase();
  if (
    normalized === "ca" ||
    normalized === "canada" ||
    normalized === "canadian"
  ) {
    return "CA";
  }
  if (normalized === "not sure" || normalized === "ns") {
    return "Not Sure";
  }
  return "US";
}

function importFieldValue(
  key: keyof ClientProfileEntry,
  raw: string,
): string | FinancialAccount[] {
  const value = raw.trim();
  if (!value || value === "—") {
    return key === "financialAccounts" ? [] : "";
  }
  if (key === "ssn" && value === "***") {
    return "";
  }
  if (key === "doi") {
    return parseImportDate(value);
  }
  if (key === "fiscalYearEnd") {
    return normalizeYearEnd(value);
  }
  if (key === "country") {
    return parseImportCountry(value);
  }
  if (key === "paymentType") {
    return normalizePaymentType(value);
  }
  if (key === "payrollReportsNeeded") {
    return normalizePayrollReportsNeeded(value);
  }
  if (key === "hasUsdAndCadAccounts" || key === "onCatchUpPlan") {
    return normalizeYesNo(value);
  }
  if (key === "meetingFrequency") {
    return normalizeClientMeetingFrequency(value);
  }
  if (key === "financialAccounts") {
    return parseFinancialAccountsFromText(value);
  }
  return value;
}

type ImportMatchMode = "replace" | "skip";

type ImportPlannedRow = {
  values: Partial<ClientProfileEntry>;
  legalName: string;
  uniqueIdRaw: string;
  existingIndex: number;
};

type ImportPreview = {
  fileName: string;
  planned: ImportPlannedRow[];
  createdCount: number;
  updatedCount: number;
};

function formatProfileCell(
  key: keyof ClientProfileEntry,
  value: unknown,
  options?: { revealSensitive?: boolean },
): string {
  if (key === "financialAccounts") {
    const accounts = normalizeFinancialAccounts(value);
    return formatFinancialAccountsList(accounts);
  }
  const text = String(value ?? "").trim();
  if (!text) {
    return "";
  }
  if (key === "ssn" && !options?.revealSensitive) {
    return "***";
  }
  if (key === "doi") {
    return formatProfileDate(text);
  }
  if (key === "paymentType") {
    if (text === "Cash") {
      return "Cash basis";
    }
    if (text === "Accrual") {
      return "Accrual basis";
    }
  }
  if (key === "payrollReportsNeeded") {
    if (text === "Yes") {
      return "Yes — need reports";
    }
    if (text === "No") {
      return "No — reports not needed";
    }
  }
  if (key === "hasUsdAndCadAccounts") {
    if (text === "Yes") {
      return "Yes — USD and CAD accounts";
    }
    if (text === "No") {
      return "No — single currency";
    }
  }
  if (key === "onCatchUpPlan") {
    if (text === "Yes") {
      return "Yes — on catch-up plan";
    }
    if (text === "No") {
      return "No — monthly cycle";
    }
  }
  return text;
}

function toWebsiteHref(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }
  return `https://${trimmed}`;
}

type ProposalCellActions = {
  onView: (profile: ClientProfileEntry) => void;
};

function renderProfileCell(
  key: keyof ClientProfileEntry,
  value: unknown,
  profile?: ClientProfileEntry,
  proposalActions?: ProposalCellActions,
): ReactNode {
  if (key === "activeProposalReference") {
    const reference =
      profile?.activeProposalReference?.trim() ||
      String(value ?? "").trim();
    const slug = profile?.activeProposalSlug?.trim() || "";
    const title =
      profile?.activeProposalName?.trim() || reference || "Active proposal";
    if (!reference && !slug) {
      return "—";
    }
    if (slug && profile && proposalActions) {
      return (
        <button
          type="button"
          className="table-link-button"
          title={title}
          onClick={() => proposalActions.onView(profile)}
        >
          {reference || "Proposal"}
        </button>
      );
    }
    return reference || "—";
  }
  const text = formatProfileCell(key, value);
  if (!text) {
    return "—";
  }
  if (key === "website" || key === "portalLink") {
    return (
      <a
        className="table-link"
        href={toWebsiteHref(text)}
        target="_blank"
        rel="noopener noreferrer"
      >
        {text}
      </a>
    );
  }
  if (key === "financialAccounts") {
    return <span className="financial-accounts-cell">{text}</span>;
  }
  return text;
}

async function persistProfiles(profiles: ClientProfileEntry[]) {
  const response = await portalFetch("/api/client-profiles", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ profiles }),
  });

  if (!response.ok) {
    const data = (await response.json()) as { error?: string };
    throw new Error(data.error || "Failed to save client profiles.");
  }
}

function profileSnapshot(entry: ClientProfileEntry) {
  return snapshotFromLabeledFields(
    TABLE_COLUMNS.map((column) => ({
      label: column.label,
      value: formatProfileCell(column.key, entry[column.key]),
    })),
  );
}

function FormSection({
  title,
  description,
  children,
}: {
  title: string;
  description?: string;
  children: ReactNode;
}) {
  return (
    <section className="client-form-section">
      <header className="client-form-section-header">
        <h4 className="client-form-section-title">{title}</h4>
        {description ? (
          <p className="client-form-section-desc">{description}</p>
        ) : null}
      </header>
      <div className="client-form-grid">{children}</div>
    </section>
  );
}

export function ClientProfilePanel() {
  const { instance, accounts } = useMsal();
  const { canEdit } = useAccess();
  const actor = getActorFromAccount(accounts[0]);
  const account = accounts[0] ?? null;
  const [profiles, setProfiles] = useState<ClientProfileEntry[]>([]);
  const [lookups, setLookups] = useState<LookupDataStore | null>(null);
  const [formOpen, setFormOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [createSharePointSite, setCreateSharePointSite] = useState(false);
  const [sharePointMemberEmails, setSharePointMemberEmails] = useState("");
  const [sharePointSites, setSharePointSites] = useState<SharePointSiteOption[]>(
    [],
  );
  const [loadingSharePointSites, setLoadingSharePointSites] = useState(false);
  const [sharePointSitesError, setSharePointSitesError] = useState<
    string | null
  >(null);
  const [form, setForm] = useState<ClientProfileInput>(
    createEmptyClientProfileForm(),
  );
  const [yearEndDay, setYearEndDay] = useState("");
  const [yearEndMonth, setYearEndMonth] = useState("");
  const [ready, setReady] = useState(false);
  const [mounted, setMounted] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [deleting, setDeleting] = useState(false);
  const [importing, setImporting] = useState(false);
  const [importPreview, setImportPreview] = useState<ImportPreview | null>(
    null,
  );
  const [proposalViewer, setProposalViewer] =
    useState<IgnitionProposalViewerTarget | null>(null);
  const [importMatchMode, setImportMatchMode] =
    useState<ImportMatchMode>("replace");
  const importInputRef = useRef<HTMLInputElement>(null);
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState("nameAsc");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(8);
  const [activeFilters, setActiveFilters] = useState<Record<string, string[]>>(
    {},
  );
  const tableWrapRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function loadProfiles() {
      try {
        const [profilesRes, lookupsRes] = await Promise.all([
          fetch("/api/client-profiles"),
          fetch("/api/lookup-data", { cache: "no-store" }),
        ]);

        const data = (await profilesRes.json()) as {
          profiles?: ClientProfileEntry[];
          error?: string;
        };
        const lookupsData = (await lookupsRes.json()) as LookupDataStore & {
          error?: string;
        };

        if (!profilesRes.ok) {
          throw new Error(data.error || "Failed to load client profiles.");
        }
        if (!lookupsRes.ok) {
          throw new Error(lookupsData.error || "Failed to load lookup data.");
        }

        if (!cancelled) {
          setProfiles(data.profiles ?? []);
          setLookups({ fields: lookupsData.fields ?? [] });
          setError(null);
        }
      } catch (loadError) {
        if (!cancelled) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load client profiles.",
          );
        }
      } finally {
        if (!cancelled) {
          setReady(true);
        }
      }
    }

    void loadProfiles();
    return () => {
      cancelled = true;
    };
  }, []);

  function syncYearEndParts(value: string) {
    const parts = parseYearEnd(value);
    setYearEndDay(parts?.day ?? "");
    setYearEndMonth(parts?.month ?? "");
  }

  function openAddForm() {
    setEditingId(null);
    setCreateSharePointSite(false);
    setSharePointMemberEmails("");
    setSharePointSitesError(null);
    setForm(
      createEmptyClientProfileForm(profiles.map((profile) => profile.uniqueId)),
    );
    syncYearEndParts("");
    setError(null);
    setFormOpen(true);
  }

  function openEditForm(profile: ClientProfileEntry) {
    setEditingId(profile.id);
    setCreateSharePointSite(false);
    setSharePointMemberEmails("");
    setSharePointSitesError(null);
    setForm({
      uniqueId: profile.uniqueId,
      legalName: profile.legalName,
      operatingName: profile.operatingName,
      website: profile.website,
      portalLink: profile.portalLink,
      industry: profile.industry,
      businessActivity: profile.businessActivity,
      entityType: profile.entityType,
      fiscalYearEnd: profile.fiscalYearEnd,
      country: profile.country,
      reportingCurrency:
        profile.reportingCurrency ||
        reportingCurrencyForCountry(profile.country),
      bd1Template: normalizeBd1Template(profile.bd1Template || "template-1"),
      serviceFrequency: profile.serviceFrequency || "",
      reportingTier: profile.reportingTier || "",
      paymentType: profile.paymentType || "",
      financialAccounts: normalizeFinancialAccounts(profile.financialAccounts),
      hasUsdAndCadAccounts: profile.hasUsdAndCadAccounts || "",
      meetingFrequency: profile.meetingFrequency || "",
      onCatchUpPlan: profile.onCatchUpPlan || "",
      payrollProvider: profile.payrollProvider || "",
      payrollReportsNeeded: profile.payrollReportsNeeded || "",
      documentRequestContact: profile.documentRequestContact || "",
      documentRequestEmail: profile.documentRequestEmail || "",
      documentRequestPhone: profile.documentRequestPhone || "",
      financialsContact: profile.financialsContact || "",
      financialsEmail: profile.financialsEmail || "",
      financialsPhone: profile.financialsPhone || "",
      state: profile.state,
      streetAddress: profile.streetAddress,
      businessAccountNumber: profile.businessAccountNumber,
      timeZone: profile.timeZone,
      ssn: profile.ssn,
      doi: profile.doi,
      ein: profile.ein,
      ctTaxRegistrationNo: profile.ctTaxRegistrationNo,
      canadianBusinessNumber: profile.canadianBusinessNumber,
      gstHst: profile.gstHst,
      provincialSalesTaxNumbers: profile.provincialSalesTaxNumbers,
      payrollAccountNumbers: profile.payrollAccountNumbers,
      ownerName: profile.ownerName,
      ownerEmail: profile.ownerEmail,
      ownerPhone: profile.ownerPhone,
      pointOfContact: profile.pointOfContact,
      pocEmail: profile.pocEmail,
      pocPhone: profile.pocPhone,
      ignitionClientSlug: profile.ignitionClientSlug || "",
      activeProposalSlug: profile.activeProposalSlug || "",
      activeProposalName: profile.activeProposalName || "",
      activeProposalReference: profile.activeProposalReference || "",
      activeProposalViewUrl: profile.activeProposalViewUrl || "",
      activeProposalPdfUrl: profile.activeProposalPdfUrl || "",
    });
    syncYearEndParts(profile.fiscalYearEnd);
    setError(null);
    setFormOpen(true);
  }

  function openActiveProposal(
    profile: ClientProfileEntry,
    mode: "view" | "download",
  ) {
    const slug = profile.activeProposalSlug.trim();
    if (!slug) {
      return;
    }
    setProposalViewer({
      slug,
      reference:
        profile.activeProposalReference.trim() ||
        profile.activeProposalName.trim() ||
        slug,
      clientLabel:
        profile.legalName.trim() ||
        profile.operatingName.trim() ||
        "Client",
      mode,
    });
  }

  function closeAddForm() {
    if (saving) {
      return;
    }
    setFormOpen(false);
    setEditingId(null);
    setCreateSharePointSite(false);
    setSharePointMemberEmails("");
    setSharePointSitesError(null);
    setError(null);
    syncYearEndParts("");
  }

  const portalLinkOptions = useMemo((): SearchableOption[] => {
    const legalName = form.legalName.trim();
    // Only CLIENT - / CLIENTS - sites (not All Company, Team Site, etc.).
    const options: SearchableOption[] = sharePointSites
      .filter(isClientSharePointSite)
      .map((site) => ({
        value: site.webUrl,
        label: site.displayName,
        meta: site.webUrl,
        keywords: `${site.displayName} ${site.webUrl}`,
      }));

    // Do not mix in saved profile links — only live CLIENT -/CLIENTS - sites.

    return options.sort((a, b) => {
      const aMatch = siteMatchesLegalName(a.label, legalName) ? 0 : 1;
      const bMatch = siteMatchesLegalName(b.label, legalName) ? 0 : 1;
      if (aMatch !== bMatch) {
        return aMatch - bMatch;
      }
      return a.label.localeCompare(b.label, undefined, { sensitivity: "base" });
    });
  }, [sharePointSites, profiles, form.legalName]);

  useEffect(() => {
    // Load existing sites for Add (picker) and Edit (link existing site).
    if (!formOpen || createSharePointSite || !account) {
      return;
    }
    let cancelled = false;
    async function loadSites() {
      setLoadingSharePointSites(true);
      setSharePointSitesError(null);
      try {
        const sites = await listSharePointSitesForPicker(
          instance,
          account!,
          form.legalName,
        );
        if (!cancelled) {
          setSharePointSites(sites);
        }
      } catch (loadError) {
        if (!cancelled) {
          setSharePointSites([]);
          setSharePointSitesError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load SharePoint sites.",
          );
        }
      } finally {
        if (!cancelled) {
          setLoadingSharePointSites(false);
        }
      }
    }
    void loadSites();
    return () => {
      cancelled = true;
    };
  }, [
    formOpen,
    createSharePointSite,
    account,
    instance,
    form.legalName,
  ]);

  function updateField<K extends keyof ClientProfileInput>(
    key: K,
    value: ClientProfileInput[K],
  ) {
    setForm((current) => ({ ...current, [key]: value }));
  }

  function addFinancialAccount() {
    setForm((current) => ({
      ...current,
      financialAccounts: [
        ...current.financialAccounts,
        createEmptyFinancialAccount(),
      ],
    }));
  }

  function updateFinancialAccount<K extends keyof FinancialAccount>(
    accountId: string,
    key: K,
    value: FinancialAccount[K],
  ) {
    setForm((current) => ({
      ...current,
      financialAccounts: current.financialAccounts.map((account) =>
        account.id === accountId ? { ...account, [key]: value } : account,
      ),
    }));
  }

  function removeFinancialAccount(accountId: string) {
    setForm((current) => ({
      ...current,
      financialAccounts: current.financialAccounts.filter(
        (account) => account.id !== accountId,
      ),
    }));
  }

  function handleServiceFrequencyChange(value: string) {
    const normalized = normalizeServiceFrequency(value);
    setForm((current) => ({
      ...current,
      serviceFrequency: normalized,
      onCatchUpPlan:
        normalized === "Catch-up/Cleanup" ? "Yes" : current.onCatchUpPlan,
    }));
  }

  function switchToMonthlyCycle() {
    setForm((current) => ({
      ...current,
      onCatchUpPlan: "No",
      serviceFrequency: "Monthly",
    }));
  }

  function setYearEndParts(day: string, month: string) {
    setYearEndDay(day);
    setYearEndMonth(month);
    if (day && month) {
      updateField(
        "fiscalYearEnd",
        formatYearEnd(Number.parseInt(day, 10), Number.parseInt(month, 10)),
      );
      return;
    }
    updateField("fiscalYearEnd", "");
  }

  function handleCountryChange(country: ClientCountry) {
    setForm((current) => {
      if (current.country === country) {
        return current;
      }

      // Exclude this profile's current ID so sequence can advance cleanly.
      const existingIds = profiles
        .filter((profile) => profile.id !== editingId)
        .map((profile) => profile.uniqueId);

      return {
        ...current,
        country,
        reportingCurrency: reportingCurrencyForCountry(country),
        state: country === "Not Sure" ? "Not Sure" : "",
        uniqueId: generateUniqueId(existingIds, country),
        // Clear country-specific tax fields that no longer apply.
        canadianBusinessNumber:
          country === "CA" ? current.canadianBusinessNumber : "",
        gstHst: country === "CA" ? current.gstHst : "",
        provincialSalesTaxNumbers:
          country === "US" ? current.provincialSalesTaxNumbers : "",
        businessAccountNumber:
          country === "CA" ? current.businessAccountNumber : "",
        ssn: country === "US" ? current.ssn : "",
        ein: country === "US" ? current.ein : "",
      };
    });
  }

  function toggleSelected(id: string) {
    setSelectedIds((current) =>
      current.includes(id)
        ? current.filter((item) => item !== id)
        : [...current, id],
    );
  }

  async function handleDeleteSelected() {
    if (selectedIds.length === 0) {
      return;
    }
    if (
      !window.confirm(
        `Delete ${selectedIds.length} selected client profile${selectedIds.length === 1 ? "" : "s"}?`,
      )
    ) {
      return;
    }

    const removed = profiles.filter((profile) =>
      selectedIds.includes(profile.id),
    );
    const next = profiles.filter(
      (profile) => !selectedIds.includes(profile.id),
    );
    setDeleting(true);
    setError(null);
    try {
      await persistProfiles(next);
      setProfiles(next);
      setSelectedIds([]);
      await Promise.all(
        removed.map((profile) =>
          logPortalAction({
            actor,
            project: profile.legalName || profile.uniqueId,
            entityType: "clientProfile",
            entityLabel: profile.legalName || profile.uniqueId,
            action: "deleted",
            summary: `Client profile deleted: ${profile.legalName || profile.uniqueId}`,
            details: Object.entries(profileSnapshot(profile))
              .filter(([, value]) => value.trim())
              .map(([label, value]) => `${label}: ${value}`)
              .join("\n"),
            comment: "Record removed from Client Profile.",
          }),
        ),
      );
    } catch (deleteError) {
      setError(
        deleteError instanceof Error
          ? deleteError.message
          : "Failed to delete client profiles.",
      );
    } finally {
      setDeleting(false);
    }
  }

  const stateOptions = getRegionsForCountry(form.country);
  const stateLabel = "State/Province";

  const industryOptions = useMemo(
    () => getLookupOptions(lookups, "industry"),
    [lookups],
  );

  const businessActivityOptions = useMemo(
    () => getLookupOptions(lookups, "businessActivity"),
    [lookups],
  );

  const timeZoneOptions = useMemo(
    () => getLookupOptions(lookups, "timeZone"),
    [lookups],
  );

  const entityTypeOptions = useMemo(
    () => getLookupOptions(lookups, "entityType"),
    [lookups],
  );

  const serviceFrequencyOptions = useMemo(
    () => getLookupOptions(lookups, "serviceFrequency"),
    [lookups],
  );

  const reportingTierOptions = useMemo(
    () => getLookupOptions(lookups, "reportingTier"),
    [lookups],
  );

  const payrollProviderOptions = useMemo(
    () => getLookupOptions(lookups, "payrollProvider"),
    [lookups],
  );

  async function createLookupValueFromProfile(
    fieldId:
      | "industry"
      | "businessActivity"
      | "timeZone"
      | "entityType"
      | "serviceFrequency"
      | "reportingTier"
      | "payrollProvider",
    formKey:
      | "industry"
      | "businessActivity"
      | "timeZone"
      | "entityType"
      | "serviceFrequency"
      | "reportingTier"
      | "payrollProvider",
    fieldLabel: string,
    label: string,
  ) {
    const trimmed = label.trim();
    if (!trimmed) {
      return;
    }
    if (!lookups) {
      updateField(formKey, trimmed);
      return;
    }

    const existing = getLookupOptions(lookups, fieldId).find(
      (option) => option.toLowerCase() === trimmed.toLowerCase(),
    );
    if (existing) {
      updateField(formKey, existing);
      return;
    }

    const hasField = lookups.fields.some((field) => field.id === fieldId);
    const next: LookupDataStore = {
      fields: hasField
        ? lookups.fields.map((field) =>
            field.id === fieldId
              ? { ...field, items: [...field.items, createLookupItem(trimmed)] }
              : field,
          )
        : [
            ...lookups.fields,
            {
              id: fieldId,
              name: fieldLabel,
              items: [createLookupItem(trimmed)],
            },
          ],
    };
    const saved = await persistLookupData(next);
    setLookups({ fields: saved.fields ?? [] });
    updateField(formKey, trimmed);
    await logPortalAction({
      actor,
      project: fieldLabel,
      entityType: "lookupData",
      entityLabel: `${fieldLabel} / ${trimmed}`,
      action: "created",
      summary: `Lookup item created: ${trimmed}`,
      details: `Field: ${fieldLabel}\nItem: ${trimmed}\nSource: Client Profile`,
      comment: "Lookup sub-item created from Client Profile dropdown.",
    });
  }

  const toolbarFilters = useMemo(
    () => [
      {
        id: "country",
        label: "Country",
        options: uniqueSortedValues(profiles.map((profile) => profile.country)),
      },
      {
        id: "state",
        label: "State/Province",
        options: uniqueSortedValues(profiles.map((profile) => profile.state)),
      },
      {
        id: "industry",
        label: "Industry",
        options: uniqueSortedValues(
          profiles.map((profile) => profile.industry),
        ),
      },
      {
        id: "businessActivity",
        label: "Business Activity",
        options: uniqueSortedValues(
          profiles.map((profile) => profile.businessActivity),
        ),
      },
      {
        id: "entityType",
        label: "Entity Type",
        options: uniqueSortedValues(
          profiles.map((profile) => profile.entityType),
        ),
      },
      {
        id: "ownerName",
        label: "Owner",
        options: uniqueSortedValues(
          profiles.map((profile) => profile.ownerName),
        ),
      },
      {
        id: "pointOfContact",
        label: "POC",
        options: uniqueSortedValues(
          profiles.map((profile) => profile.pointOfContact),
        ),
      },
    ],
    [profiles],
  );

  const visibleProfiles = useMemo(() => {
    const query = search.trim().toLowerCase();
    let rows = profiles.filter((profile) => {
      if (query) {
        const haystack = [profile.uniqueId, profile.legalName]
          .join(" ")
          .toLowerCase();
        if (!haystack.includes(query)) {
          return false;
        }
      }

      for (const [filterId, selected] of Object.entries(activeFilters)) {
        if (!selected.length) {
          continue;
        }
        const value = String(
          profile[filterId as keyof ClientProfileEntry] ?? "",
        ).trim();
        if (!selected.includes(value)) {
          return false;
        }
      }
      return true;
    });

    rows = [...rows].sort((a, b) => {
      switch (sort) {
        case "idAsc":
          return a.uniqueId.localeCompare(b.uniqueId, undefined, {
            sensitivity: "base",
          });
        case "newest":
          return (
            new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
          );
        case "oldest":
          return (
            new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime()
          );
        case "nameAsc":
        default:
          return a.legalName.localeCompare(b.legalName, undefined, {
            sensitivity: "base",
          });
      }
    });

    return rows;
  }, [profiles, search, activeFilters, sort]);

  const totalPages = Math.max(
    1,
    Math.ceil(visibleProfiles.length / Math.max(pageSize, MIN_PAGE_SIZE)),
  );
  const currentPage = Math.min(page, totalPages);

  const pageProfiles = useMemo(() => {
    const size = Math.max(pageSize, MIN_PAGE_SIZE);
    const start = (currentPage - 1) * size;
    return visibleProfiles.slice(start, start + size);
  }, [visibleProfiles, currentPage, pageSize]);

  useEffect(() => {
    setPage(1);
  }, [search, sort, activeFilters]);

  useEffect(() => {
    if (page > totalPages) {
      setPage(totalPages);
    }
  }, [page, totalPages]);

  useEffect(() => {
    if (!ready) {
      return;
    }
    const wrap = tableWrapRef.current;
    if (!wrap) {
      return;
    }

    let frame = 0;
    let lastAppliedSize = 0;

    function measurePageSize() {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => {
        const container = tableWrapRef.current;
        if (!container) {
          return;
        }

        const available = container.clientHeight;
        if (available <= 0) {
          return;
        }

        // Use fixed row height so measuring content cannot feed back into
        // pageSize → row count → ResizeObserver → measure (screen jerk).
        const header =
          container.querySelector("thead")?.getBoundingClientRect().height ??
          FALLBACK_HEADER_HEIGHT;
        const usable = Math.max(0, available - header);
        const nextSize = Math.max(
          MIN_PAGE_SIZE,
          Math.floor(usable / FALLBACK_ROW_HEIGHT),
        );
        if (nextSize === lastAppliedSize) {
          return;
        }
        lastAppliedSize = nextSize;
        setPageSize((current) => (current === nextSize ? current : nextSize));
      });
    }

    measurePageSize();
    const observer = new ResizeObserver(() => {
      measurePageSize();
    });
    observer.observe(wrap);
    window.addEventListener("resize", measurePageSize);

    return () => {
      cancelAnimationFrame(frame);
      observer.disconnect();
      window.removeEventListener("resize", measurePageSize);
    };
  }, [ready]);

  function handleFilterChange(filterId: string, values: string[]) {
    setActiveFilters((current) => ({ ...current, [filterId]: values }));
  }

  function toggleSelectAll() {
    const pageIds = pageProfiles.map((profile) => profile.id);
    const allPageSelected =
      pageIds.length > 0 && pageIds.every((id) => selectedIds.includes(id));
    if (allPageSelected) {
      setSelectedIds((current) =>
        current.filter((id) => !pageIds.includes(id)),
      );
      return;
    }
    setSelectedIds((current) => [...new Set([...current, ...pageIds])]);
  }

  function handleExport() {
    const header = TABLE_COLUMNS.map((column) => column.label);
    const rows = visibleProfiles.map((profile) =>
      TABLE_COLUMNS.map((column) =>
        formatProfileCell(column.key, profile[column.key]),
      ),
    );
    downloadCsv("client-profiles.csv", [header, ...rows]);
  }

  function openImportPicker() {
    importInputRef.current?.click();
  }

  function closeImportPreview() {
    setImportPreview(null);
    setImportMatchMode("replace");
    setImporting(false);
    if (importInputRef.current) {
      importInputRef.current.value = "";
    }
  }

  async function applyImport(
    preview: ImportPreview,
    matchMode: ImportMatchMode,
  ) {
    const next = [...profiles];
    const usedUniqueIds = next.map((profile) => profile.uniqueId);
    let created = 0;
    let updated = 0;
    let skipped = 0;

    for (const row of preview.planned) {
      if (row.existingIndex >= 0) {
        if (matchMode === "skip") {
          skipped += 1;
          continue;
        }
        const existing = next[row.existingIndex];
        next[row.existingIndex] = normalizeClientProfile({
          ...existing,
          ...row.values,
          id: existing.id,
          uniqueId: existing.uniqueId,
          legalName: row.legalName,
          ssn: row.values.ssn?.trim() ? row.values.ssn : existing.ssn,
          createdAt: existing.createdAt,
        });
        updated += 1;
        continue;
      }

      const country = parseImportCountry(String(row.values.country ?? "US"));
      const uniqueId =
        row.uniqueIdRaw || generateUniqueId(usedUniqueIds, country);
      usedUniqueIds.push(uniqueId);
      next.unshift(
        normalizeClientProfile({
          id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          ...row.values,
          uniqueId,
          legalName: row.legalName,
          country,
          createdAt: new Date().toISOString(),
        }),
      );
      created += 1;
    }

    if (created === 0 && updated === 0) {
      throw new Error(
        skipped > 0
          ? "All matching rows were skipped. Nothing was imported."
          : "No valid rows found. Legal Name is required.",
      );
    }

    await persistProfiles(next);
    setProfiles(next);
    setSelectedIds([]);
    setPage(1);

    const parts = [`${created} added`, `${updated} updated`];
    if (skipped > 0) {
      parts.push(`${skipped} skipped`);
    }
    setNotice(`Import complete: ${parts.join(", ")}.`);

    await logPortalAction({
      actor,
      project: "Client Profile",
      entityType: "clientProfile",
      entityLabel: "CSV Import",
      action: "updated",
      summary: `Imported client profiles (${created} added, ${updated} updated${
        skipped > 0 ? `, ${skipped} skipped` : ""
      })`,
      details: `File: ${preview.fileName}\nAdded: ${created}\nUpdated: ${updated}\nSkipped: ${skipped}\nMatch mode: ${matchMode}`,
      comment: "Bulk import from CSV.",
    });
  }

  async function handleImportFile(file: File) {
    setImporting(true);
    setError(null);
    setNotice(null);
    setImportPreview(null);
    try {
      const text = await file.text();
      const rows = parseCsv(text);
      if (rows.length < 2) {
        throw new Error(
          "CSV must include a header row and at least one data row.",
        );
      }

      const headers = rows[0].map((header) => header.trim().toLowerCase());
      const labelToKey = new Map(
        TABLE_COLUMNS.map((column) => [
          column.label.trim().toLowerCase(),
          column.key,
        ]),
      );
      for (const column of TABLE_COLUMNS) {
        labelToKey.set(column.key.toLowerCase(), column.key);
      }

      const columnIndexes: { key: keyof ClientProfileEntry; index: number }[] =
        [];
      headers.forEach((header, index) => {
        const key = labelToKey.get(header);
        // Ignition proposal links are filled by Sync — skip on CSV import.
        if (
          key &&
          key !== "ignitionClientSlug" &&
          !String(key).startsWith("activeProposal")
        ) {
          columnIndexes.push({ key, index });
        }
      });

      if (columnIndexes.length === 0) {
        throw new Error(
          "No recognized columns. Export a CSV first and use those headers.",
        );
      }

      const planned: ImportPlannedRow[] = [];
      const matchedIndexes = new Set<number>();

      for (const row of rows.slice(1)) {
        const values: Partial<ClientProfileEntry> = {};
        for (const { key, index } of columnIndexes) {
          const imported = importFieldValue(key, row[index] ?? "");
          if (key === "ssn" && !imported) {
            continue;
          }
          if (key === "financialAccounts") {
            values.financialAccounts = Array.isArray(imported)
              ? imported
              : [];
            continue;
          }
          (values as Record<string, string>)[key] = String(imported);
        }

        const legalName = String(values.legalName ?? "").trim();
        if (!legalName) {
          continue;
        }

        const uniqueIdRaw = String(values.uniqueId ?? "").trim();
        const existingIndex = profiles.findIndex((profile, index) => {
          if (matchedIndexes.has(index)) {
            return false;
          }
          if (
            uniqueIdRaw &&
            profile.uniqueId.trim().toLowerCase() === uniqueIdRaw.toLowerCase()
          ) {
            return true;
          }
          return (
            profile.legalName.trim().toLowerCase() === legalName.toLowerCase()
          );
        });

        if (existingIndex >= 0) {
          matchedIndexes.add(existingIndex);
        }

        planned.push({ values, legalName, uniqueIdRaw, existingIndex });
      }

      const createdCount = planned.filter((row) => row.existingIndex < 0).length;
      const updatedCount = planned.filter(
        (row) => row.existingIndex >= 0,
      ).length;

      if (createdCount === 0 && updatedCount === 0) {
        throw new Error("No valid rows found. Legal Name is required.");
      }

      const preview: ImportPreview = {
        fileName: file.name,
        planned,
        createdCount,
        updatedCount,
      };

      if (updatedCount > 0) {
        setImportMatchMode("replace");
        setImportPreview(preview);
        return;
      }

      await applyImport(preview, "replace");
    } catch (importError) {
      setError(
        importError instanceof Error
          ? importError.message
          : "Failed to import client profiles.",
      );
    } finally {
      setImporting(false);
      if (importInputRef.current) {
        importInputRef.current.value = "";
      }
    }
  }

  async function confirmImportPreview() {
    if (!importPreview) {
      return;
    }
    setImporting(true);
    setError(null);
    try {
      await applyImport(importPreview, importMatchMode);
      closeImportPreview();
    } catch (importError) {
      setError(
        importError instanceof Error
          ? importError.message
          : "Failed to import client profiles.",
      );
      closeImportPreview();
    }
  }

  async function handleSave(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    const legalName = form.legalName.trim();
    if (!legalName) {
      setError("Legal Name is required.");
      return;
    }

    const duplicateLegalName = profiles.some(
      (profile) =>
        profile.id !== editingId &&
        profile.legalName.trim().toLowerCase() === legalName.toLowerCase(),
    );
    if (duplicateLegalName) {
      setError(
        `Legal Name "${legalName}" already exists. Choose a different name.`,
      );
      return;
    }

    const existing = editingId
      ? profiles.find((profile) => profile.id === editingId)
      : null;

    const uniqueId =
      form.uniqueId.trim() ||
      generateUniqueId(
        profiles
          .filter((profile) => profile.id !== editingId)
          .map((profile) => profile.uniqueId),
        form.country,
      );

    let portalLink = form.portalLink.trim();
    let memberWarnings: string[] = [];
    if (!editingId && createSharePointSite) {
      if (!account) {
        setError("Sign in with Microsoft to create a SharePoint site.");
        return;
      }
      setSaving(true);
      setError(null);
      try {
        const created = await createClientSharePointSite(
          instance,
          account,
          legalName,
          parseMemberEmails(sharePointMemberEmails),
        );
        portalLink = created.webUrl;
        memberWarnings = created.memberWarnings;
        setForm((current) => ({ ...current, portalLink }));
      } catch (createError) {
        setSaving(false);
        setError(
          createError instanceof Error
            ? createError.message
            : "Failed to create SharePoint site.",
        );
        return;
      }
    }

    const entry: ClientProfileEntry = {
      id: existing?.id ?? `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      ...form,
      uniqueId,
      legalName,
      operatingName: form.operatingName.trim(),
      website: form.website.trim(),
      portalLink,
      industry: form.industry.trim(),
      businessActivity: form.businessActivity.trim(),
      entityType: form.entityType.trim(),
      fiscalYearEnd: normalizeYearEnd(form.fiscalYearEnd),
      reportingCurrency: reportingCurrencyForCountry(form.country),
      bd1Template: normalizeBd1Template(form.bd1Template),
      serviceFrequency: normalizeServiceFrequency(form.serviceFrequency),
      reportingTier: normalizeReportingTier(form.reportingTier),
      paymentType: normalizePaymentType(form.paymentType),
      financialAccounts: normalizeFinancialAccounts(form.financialAccounts),
      hasUsdAndCadAccounts: normalizeYesNo(form.hasUsdAndCadAccounts),
      meetingFrequency: normalizeClientMeetingFrequency(form.meetingFrequency),
      onCatchUpPlan: normalizeYesNo(form.onCatchUpPlan),
      state: form.state.trim(),
      streetAddress: form.streetAddress.trim(),
      businessAccountNumber:
        form.country === "CA" ? form.businessAccountNumber.trim() : "",
      timeZone: form.timeZone.trim(),
      ssn: form.country === "US" ? form.ssn.trim() : "",
      doi: form.doi.trim(),
      ein: form.country === "US" ? form.ein.trim() : "",
      ctTaxRegistrationNo: form.ctTaxRegistrationNo.trim(),
      canadianBusinessNumber:
        form.country === "CA" ? form.canadianBusinessNumber.trim() : "",
      gstHst: form.country === "CA" ? form.gstHst.trim() : "",
      provincialSalesTaxNumbers:
        form.country === "US" ? form.provincialSalesTaxNumbers.trim() : "",
      payrollProvider: form.payrollProvider.trim(),
      payrollReportsNeeded: normalizePayrollReportsNeeded(
        form.payrollReportsNeeded,
      ),
      payrollAccountNumbers: form.payrollAccountNumbers.trim(),
      ownerName: form.ownerName.trim(),
      ownerEmail: form.ownerEmail.trim(),
      ownerPhone: form.ownerPhone.trim(),
      pointOfContact: form.pointOfContact.trim(),
      pocEmail: form.pocEmail.trim(),
      pocPhone: form.pocPhone.trim(),
      documentRequestContact: form.documentRequestContact.trim(),
      documentRequestEmail: form.documentRequestEmail.trim(),
      documentRequestPhone: form.documentRequestPhone.trim(),
      financialsContact: form.financialsContact.trim(),
      financialsEmail: form.financialsEmail.trim(),
      financialsPhone: form.financialsPhone.trim(),
      // Preserve Ignition proposal links from sync (not edited in this form).
      ignitionClientSlug: existing?.ignitionClientSlug ?? "",
      activeProposalSlug: existing?.activeProposalSlug ?? "",
      activeProposalName: existing?.activeProposalName ?? "",
      activeProposalReference: existing?.activeProposalReference ?? "",
      activeProposalViewUrl: existing?.activeProposalViewUrl ?? "",
      activeProposalPdfUrl: existing?.activeProposalPdfUrl ?? "",
      createdAt: existing?.createdAt ?? new Date().toISOString(),
    };

    const next = editingId
      ? profiles.map((profile) => (profile.id === editingId ? entry : profile))
      : [entry, ...profiles];
    setSaving(true);
    setError(null);

    try {
      await persistProfiles(next);
      setProfiles(next);
      setFormOpen(false);
      setEditingId(null);
      setCreateSharePointSite(false);
      setSharePointMemberEmails("");
      setForm(
        createEmptyClientProfileForm(
          next.map((profile) => profile.uniqueId),
        ),
      );
      syncYearEndParts("");
      if (memberWarnings.length > 0) {
        setNotice(
          `Client saved. Some SharePoint members could not be added: ${memberWarnings.join("; ")}`,
        );
      }

      const after = profileSnapshot(entry);
      if (existing) {
        const details = buildChangeDetails(profileSnapshot(existing), after);
        await logPortalAction({
          actor,
          project: entry.legalName,
          entityType: "clientProfile",
          entityLabel: entry.legalName,
          action: "updated",
          summary: `Client profile updated: ${entry.legalName}`,
          details: details || "No field values changed.",
          comment: "Changes saved on Client Profile.",
        });
      } else {
        await logPortalAction({
          actor,
          project: entry.legalName,
          entityType: "clientProfile",
          entityLabel: entry.legalName,
          action: "created",
          summary: `Client profile created: ${entry.legalName}`,
          details: Object.entries(after)
            .filter(([, value]) => value.trim())
            .map(([label, value]) => `${label}: ${value}`)
            .join("\n"),
          comment: "New client profile added.",
        });
      }
    } catch (saveError) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save client profile.",
      );
    } finally {
      setSaving(false);
    }
  }

  const pageSelectedCount = pageProfiles.filter((profile) =>
    selectedIds.includes(profile.id),
  ).length;
  const size = Math.max(pageSize, MIN_PAGE_SIZE);
  const rangeStart =
    visibleProfiles.length === 0 ? 0 : (currentPage - 1) * size + 1;
  const rangeEnd = Math.min(currentPage * size, visibleProfiles.length);

  return (
    <section className="content-panel content-panel--actions">
      <DataListToolbar
        search={search}
        onSearchChange={(value) => {
          setSearch(value);
          setPage(1);
        }}
        searchPlaceholder="ID or Legal Name"
        sort={sort}
        onSortChange={(value) => {
          setSort(value);
          setPage(1);
        }}
        sortOptions={PROFILE_SORT_OPTIONS}
        filters={toolbarFilters}
        activeFilters={activeFilters}
        onFilterChange={handleFilterChange}
        actions={
          <>
            {canEdit ? (
              <button
                type="button"
                className="button danger"
                onClick={() => void handleDeleteSelected()}
                disabled={selectedIds.length === 0 || deleting}
              >
                {deleting ? "Deleting…" : "Delete"}
              </button>
            ) : null}
            <button
              type="button"
              className="button secondary-solid"
              onClick={handleExport}
              disabled={visibleProfiles.length === 0}
            >
              Export
            </button>
            {canEdit ? (
              <>
                <button
                  type="button"
                  className="button secondary-solid"
                  onClick={openImportPicker}
                  disabled={importing}
                >
                  {importing ? "Importing…" : "Import"}
                </button>
                <input
                  ref={importInputRef}
                  type="file"
                  accept=".csv,text/csv"
                  hidden
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (file) {
                      void handleImportFile(file);
                    }
                  }}
                />
                <button
                  type="button"
                  className="button primary"
                  onClick={openAddForm}
                >
                  Add
                </button>
              </>
            ) : null}
          </>
        }
      />

      {error && !formOpen ? (
        <div className="inline-notice inline-notice--error" role="alert">
          <p className="inline-notice__text">{error}</p>
          <button
            type="button"
            className="inline-notice__close"
            aria-label="Dismiss error"
            onClick={() => setError(null)}
          >
            ×
          </button>
        </div>
      ) : null}
      {notice && !formOpen ? (
        <div className="inline-notice inline-notice--success" role="status">
          <p className="inline-notice__text">{notice}</p>
          <button
            type="button"
            className="inline-notice__close"
            aria-label="Dismiss notification"
            onClick={() => setNotice(null)}
          >
            ×
          </button>
        </div>
      ) : null}

      {importPreview && mounted
        ? createPortal(
            <div className="dialog-backdrop" role="presentation">
              <div
                className="dialog-panel"
                role="dialog"
                aria-modal="true"
                aria-labelledby="import-match-dialog-title"
              >
                <div className="dialog-header">
                  <h3 id="import-match-dialog-title" className="dialog-title">
                    Matching rows found
                  </h3>
                  <button
                    type="button"
                    className="dialog-close"
                    aria-label="Close"
                    onClick={closeImportPreview}
                    disabled={importing}
                  >
                    ×
                  </button>
                </div>

                <p className="form-message" style={{ marginBottom: "0.85rem" }}>
                  {importPreview.updatedCount === 1
                    ? "1 row matches an existing profile"
                    : `${importPreview.updatedCount} rows match existing profiles`}{" "}
                  by Unique ID or Legal Name
                  {importPreview.createdCount > 0
                    ? `, and ${
                        importPreview.createdCount === 1
                          ? "1 new row"
                          : `${importPreview.createdCount} new rows`
                      } will be added`
                    : ""}
                  .
                </p>

                <div
                  className="radio-group"
                  role="radiogroup"
                  aria-label="How to handle matching rows"
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: "0.55rem",
                    marginBottom: "1.1rem",
                  }}
                >
                  <label className="radio-option">
                    <input
                      type="radio"
                      name="import-match-mode"
                      checked={importMatchMode === "replace"}
                      onChange={() => setImportMatchMode("replace")}
                      disabled={importing}
                    />
                    <span>Replace matching profiles</span>
                  </label>
                  <label className="radio-option">
                    <input
                      type="radio"
                      name="import-match-mode"
                      checked={importMatchMode === "skip"}
                      onChange={() => setImportMatchMode("skip")}
                      disabled={importing}
                    />
                    <span>Skip matches — only add new rows</span>
                  </label>
                </div>

                {importMatchMode === "skip" &&
                importPreview.createdCount === 0 ? (
                  <p className="form-message error" style={{ marginBottom: "0.85rem" }}>
                    There are no new rows to add if matches are skipped.
                  </p>
                ) : null}

                <div className="dialog-actions">
                  <button
                    type="button"
                    className="button secondary-solid"
                    onClick={closeImportPreview}
                    disabled={importing}
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    className="button primary"
                    onClick={() => void confirmImportPreview()}
                    disabled={
                      importing ||
                      (importMatchMode === "skip" &&
                        importPreview.createdCount === 0)
                    }
                  >
                    {importing ? "Importing…" : "Import"}
                  </button>
                </div>
              </div>
            </div>,
            document.body,
          )
        : null}

      {formOpen && mounted
        ? createPortal(
            <div className="dialog-backdrop" role="presentation">
              <div
                className="dialog-panel dialog-panel--wide"
                role="dialog"
                aria-modal="true"
                aria-labelledby="client-profile-dialog-title"
              >
                <div className="dialog-header">
                  <h3 id="client-profile-dialog-title" className="dialog-title">
                    {editingId ? "Edit Client Profile" : "Add Client Profile"}
                  </h3>
                  <button
                    type="button"
                    className="dialog-close"
                    aria-label="Close"
                    onClick={closeAddForm}
                    disabled={saving}
                  >
                    ×
                  </button>
                </div>

                {error ? <p className="form-message error">{error}</p> : null}

                <form
                  className="client-form client-form--dialog"
                  onSubmit={handleSave}
                >
          <div className="client-form-sections">
          <FormSection
            title="Basics"
            description="Who the client is and how the business is classified."
          >
            <label className="field">
              <span className="field-label">Unique ID</span>
              <input
                className="field-input field-input--readonly"
                value={form.uniqueId}
                readOnly
                tabIndex={-1}
              />
            </label>

            <label className="field">
              <span className="field-label">
                Legal Name <span className="field-required">*</span>
              </span>
              <input
                className="field-input"
                value={form.legalName}
                onChange={(e) => {
                  updateField("legalName", e.target.value);
                  if (error) {
                    setError(null);
                  }
                }}
                required
              />
            </label>

            <label className="field">
              <span className="field-label">Operating Name</span>
              <input
                className="field-input"
                value={form.operatingName}
                onChange={(e) => updateField("operatingName", e.target.value)}
              />
            </label>

            <SearchableSelect
              label="Industry"
              value={form.industry}
              options={industryOptions}
              allowCreate
              placeholder="Search/Select from the list"
              onChange={(value) => updateField("industry", value)}
              onCreateOption={(label) => {
                void createLookupValueFromProfile(
                  "industry",
                  "industry",
                  "Industry",
                  label,
                );
              }}
            />

            <SearchableSelect
              label="Business Activity"
              value={form.businessActivity}
              options={businessActivityOptions}
              allowCreate
              placeholder="Search/Select from the list"
              onChange={(value) => updateField("businessActivity", value)}
              onCreateOption={(label) => {
                void createLookupValueFromProfile(
                  "businessActivity",
                  "businessActivity",
                  "Business Activity",
                  label,
                );
              }}
            />

            <SearchableSelect
              label="Entity Type"
              value={form.entityType}
              options={entityTypeOptions}
              allowCreate
              placeholder="Search/Select from the list"
              onChange={(value) => updateField("entityType", value)}
              onCreateOption={(label) => {
                void createLookupValueFromProfile(
                  "entityType",
                  "entityType",
                  "Entity Type",
                  label,
                );
              }}
            />

            <label className="field">
              <span className="field-label">Website</span>
              <input
                className="field-input"
                type="text"
                inputMode="url"
                value={form.website}
                placeholder="https://"
                onChange={(e) => updateField("website", e.target.value)}
              />
            </label>

            <label className="field">
              <span className="field-label">Year End</span>
              <div className="year-end-row">
                <select
                  className="field-input field-select year-end-day"
                  aria-label="Year End day"
                  value={yearEndDay}
                  onChange={(e) =>
                    setYearEndParts(e.target.value, yearEndMonth)
                  }
                >
                  <option value="">Day</option>
                  {YEAR_END_DAYS.map((day) => (
                    <option key={day.value} value={day.value}>
                      {day.label}
                    </option>
                  ))}
                </select>
                <select
                  className="field-input field-select year-end-month"
                  aria-label="Year End month"
                  value={yearEndMonth}
                  onChange={(e) =>
                    setYearEndParts(yearEndDay, e.target.value)
                  }
                >
                  <option value="">Month</option>
                  {YEAR_END_MONTHS.map((month) => (
                    <option key={month.value} value={month.value}>
                      {month.label}
                    </option>
                  ))}
                </select>
              </div>
            </label>
          </FormSection>

          <FormSection
            title="Location & currency"
            description="Where the client is based and how currency is handled."
          >
            <fieldset className="field field--radio">
              <legend className="field-label">Country</legend>
              <div className="radio-row">
                {(["US", "CA", "Not Sure"] as ClientCountry[]).map((option) => (
                  <label key={option} className="radio-option">
                    <input
                      type="radio"
                      name="country"
                      value={option}
                      checked={form.country === option}
                      onChange={() => handleCountryChange(option)}
                    />
                    {option}
                  </label>
                ))}
              </div>
            </fieldset>

            <fieldset className="field field--radio">
              <legend className="field-label">Payment Type</legend>
              <div className="radio-row">
                {PAYMENT_TYPE_OPTIONS.map((option) => (
                  <label key={option} className="radio-option">
                    <input
                      type="radio"
                      name="paymentType"
                      value={option}
                      checked={form.paymentType === option}
                      onChange={() =>
                        updateField("paymentType", option as PaymentType)
                      }
                    />
                    {option === "Cash" ? "Cash basis" : "Accrual basis"}
                  </label>
                ))}
              </div>
            </fieldset>

            <label className="field">
              <span className="field-label">Reporting Currency</span>
              <input
                className="field-input"
                value={
                  form.reportingCurrency ||
                  reportingCurrencyForCountry(form.country)
                }
                readOnly
              />
            </label>

            <SearchableSelect
              label={stateLabel}
              value={form.state}
              options={stateOptions}
              placeholder={`Search ${stateLabel.toLowerCase()}`}
              onChange={(value) => updateField("state", value)}
            />

            <label className="field field--wide">
              <span className="field-label">Business Address</span>
              <input
                className="field-input"
                value={form.streetAddress}
                onChange={(e) => updateField("streetAddress", e.target.value)}
              />
            </label>

            <SearchableSelect
              label="Time Zone"
              value={form.timeZone}
              options={timeZoneOptions}
              allowCreate
              placeholder="Search/Select from the list"
              onChange={(value) => updateField("timeZone", value)}
              onCreateOption={(label) => {
                void createLookupValueFromProfile(
                  "timeZone",
                  "timeZone",
                  "Time Zone",
                  label,
                );
              }}
            />
          </FormSection>

          <FormSection
            title="Engagement"
            description="How often you work with them, reporting expectations, and plan type."
          >
            <fieldset className="field field--radio">
              <legend className="field-label">BD1 Template</legend>
              <div className="radio-row">
                {BD1_TEMPLATE_OPTIONS.map((option) => (
                  <label key={option.id} className="radio-option">
                    <input
                      type="radio"
                      name="bd1Template"
                      value={option.id}
                      checked={form.bd1Template === option.id}
                      onChange={() =>
                        updateField("bd1Template", option.id as Bd1TemplateId)
                      }
                    />
                    {option.label}
                  </label>
                ))}
              </div>
              <p className="lookup-hint">
                Template 1 uses Bank and Credit Card Statements Request.
                Template 2 uses Bank, Credit Card And Other Statements Request.
              </p>
            </fieldset>

            <SearchableSelect
              label="Service Frequency"
              value={form.serviceFrequency}
              options={serviceFrequencyOptions}
              allowCreate
              placeholder="Search/Select from the list"
              onChange={handleServiceFrequencyChange}
              onCreateOption={(label) => {
                void createLookupValueFromProfile(
                  "serviceFrequency",
                  "serviceFrequency",
                  "Service Frequency",
                  label,
                ).then(() => handleServiceFrequencyChange(label));
              }}
            />

            <SearchableSelect
              label="Reporting Tier"
              value={form.reportingTier}
              options={reportingTierOptions}
              allowCreate
              placeholder="Search/Select from the list"
              onChange={(value) => updateField("reportingTier", value)}
              onCreateOption={(label) => {
                void createLookupValueFromProfile(
                  "reportingTier",
                  "reportingTier",
                  "Reporting Tier",
                  label,
                );
              }}
            />

            <SearchableSelect
              label="Meeting Frequency"
              value={form.meetingFrequency}
              options={[...CLIENT_MEETING_FREQUENCY_OPTIONS]}
              placeholder="Select monthly, quarterly, or none"
              onChange={(value) =>
                updateField(
                  "meetingFrequency",
                  value as ClientMeetingFrequency,
                )
              }
            />

            <fieldset className="field field--radio">
              <legend className="field-label">
                Accounts in both USD and CAD?
              </legend>
              <div className="radio-row">
                {YES_NO_OPTIONS.map((option) => (
                  <label key={option} className="radio-option">
                    <input
                      type="radio"
                      name="hasUsdAndCadAccounts"
                      value={option}
                      checked={form.hasUsdAndCadAccounts === option}
                      onChange={() =>
                        updateField("hasUsdAndCadAccounts", option as YesNo)
                      }
                    />
                    {option}
                  </label>
                ))}
              </div>
              <p className="lookup-hint">
                Yes includes foreign-currency accounts in the request list and
                reconciliations.
              </p>
            </fieldset>

            <div className="field field--wide catch-up-plan-field">
              <fieldset className="field field--radio">
                <legend className="field-label">On catch-up project plan?</legend>
                <div className="radio-row">
                  {YES_NO_OPTIONS.map((option) => (
                    <label key={option} className="radio-option">
                      <input
                        type="radio"
                        name="onCatchUpPlan"
                        value={option}
                        checked={form.onCatchUpPlan === option}
                        onChange={() =>
                          updateField("onCatchUpPlan", option as YesNo)
                        }
                      />
                      {option}
                    </label>
                  ))}
                </div>
              </fieldset>
              {form.onCatchUpPlan === "Yes" ? (
                <div className="catch-up-plan-actions">
                  <p className="lookup-hint">
                    Catch-up clients stay on a separate project plan until the
                    books are current.
                  </p>
                  <button
                    type="button"
                    className="button primary"
                    onClick={switchToMonthlyCycle}
                  >
                    Switch to monthly cycle
                  </button>
                </div>
              ) : null}
            </div>
          </FormSection>

          <FormSection
            title="Financial accounts"
            description="Bank, credit card, loan, and merchant accounts — and whether each needs a statement."
          >
            <div className="field field--wide financial-accounts">
              <div className="financial-accounts-header">
                <span className="field-label">Accounts</span>
                <button
                  type="button"
                  className="button secondary-solid"
                  onClick={addFinancialAccount}
                >
                  Add account
                </button>
              </div>
              {form.financialAccounts.length === 0 ? (
                <p className="lookup-hint">
                  Add bank, credit card, loan, and merchant accounts. For each,
                  note whether it is on a bank feed or needs a statement from
                  the client.
                </p>
              ) : (
                <div className="financial-accounts-list">
                  {form.financialAccounts.map((account, index) => (
                    <div
                      key={account.id}
                      className="financial-account-row"
                    >
                      <div className="financial-account-row-title">
                        <span>Account {index + 1}</span>
                        <button
                          type="button"
                          className="button secondary"
                          onClick={() => removeFinancialAccount(account.id)}
                        >
                          Remove
                        </button>
                      </div>
                      <div className="financial-account-row-grid">
                        <label className="field">
                          <span className="field-label">Account type</span>
                          <select
                            className="field-input"
                            value={account.accountType}
                            onChange={(event) =>
                              updateFinancialAccount(
                                account.id,
                                "accountType",
                                event.target.value as FinancialAccountType | "",
                              )
                            }
                          >
                            <option value="">Select type</option>
                            {FINANCIAL_ACCOUNT_TYPE_OPTIONS.map((option) => (
                              <option key={option} value={option}>
                                {option}
                              </option>
                            ))}
                          </select>
                        </label>
                        <label className="field">
                          <span className="field-label">Institution name</span>
                          <input
                            className="field-input"
                            value={account.institutionName}
                            placeholder="e.g. Chase, Amex, Square"
                            onChange={(event) =>
                              updateFinancialAccount(
                                account.id,
                                "institutionName",
                                event.target.value,
                              )
                            }
                          />
                        </label>
                        <label className="field">
                          <span className="field-label">Last 4 digits</span>
                          <input
                            className="field-input"
                            value={account.lastFour}
                            inputMode="numeric"
                            maxLength={4}
                            placeholder="1234"
                            onChange={(event) =>
                              updateFinancialAccount(
                                account.id,
                                "lastFour",
                                event.target.value.replace(/\D/g, "").slice(0, 4),
                              )
                            }
                          />
                        </label>
                        <fieldset className="field field--radio">
                          <legend className="field-label">Statement source</legend>
                          <div className="radio-row">
                            {FINANCIAL_ACCOUNT_SOURCE_OPTIONS.map((option) => (
                              <label key={option} className="radio-option">
                                <input
                                  type="radio"
                                  name={`financial-source-${account.id}`}
                                  value={option}
                                  checked={account.source === option}
                                  onChange={() =>
                                    updateFinancialAccount(
                                      account.id,
                                      "source",
                                      option as FinancialAccountSource,
                                    )
                                  }
                                />
                                {option === "Needs statement"
                                  ? "Needs statement from client"
                                  : option}
                              </label>
                            ))}
                          </div>
                        </fieldset>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </FormSection>

          <FormSection
            title="Tax & registration"
            description="IDs and numbers used for filings and setup."
          >
            {form.country === "CA" ? (
              <label className="field">
                <span className="field-label">Business Account Number</span>
                <input
                  className="field-input"
                  value={form.businessAccountNumber}
                  onChange={(e) =>
                    updateField("businessAccountNumber", e.target.value)
                  }
                />
              </label>
            ) : null}

            {form.country === "US" ? (
              <label className="field">
                <span className="field-label">SSN</span>
                <input
                  className="field-input"
                  type="text"
                  value={form.ssn}
                  autoComplete="off"
                  onChange={(e) => updateField("ssn", e.target.value)}
                />
              </label>
            ) : null}

            <label className="field">
              <span className="field-label">DOI (MM/DD/YYYY)</span>
              <input
                className="field-input"
                type="date"
                value={form.doi}
                onChange={(e) => updateField("doi", e.target.value)}
              />
            </label>

            {form.country === "US" ? (
              <label className="field">
                <span className="field-label">FEIN</span>
                <input
                  className="field-input"
                  value={form.ein}
                  onChange={(e) => updateField("ein", e.target.value)}
                />
              </label>
            ) : null}

            <label className="field">
              <span className="field-label">Provincial Sales Tax Numbers</span>
              <input
                className="field-input"
                value={form.ctTaxRegistrationNo}
                onChange={(e) =>
                  updateField("ctTaxRegistrationNo", e.target.value)
                }
              />
            </label>

            {form.country === "CA" ? (
              <label className="field">
                <span className="field-label">Canadian Business Number</span>
                <input
                  className="field-input"
                  value={form.canadianBusinessNumber}
                  onChange={(e) =>
                    updateField("canadianBusinessNumber", e.target.value)
                  }
                />
              </label>
            ) : null}

            {form.country === "CA" ? (
              <label className="field">
                <span className="field-label">GST/HST</span>
                <input
                  className="field-input"
                  value={form.gstHst}
                  onChange={(e) => updateField("gstHst", e.target.value)}
                />
              </label>
            ) : null}
          </FormSection>

          <FormSection
            title="Portal"
            description="SharePoint site used for document uploads."
          >
            {!editingId ? (
              <label className="field field--checkbox">
                <span className="field-label">
                  <input
                    type="checkbox"
                    checked={createSharePointSite}
                    disabled={saving}
                    onChange={(event) => {
                      const checked = event.target.checked;
                      setCreateSharePointSite(checked);
                      if (!checked) {
                        setSharePointMemberEmails("");
                      } else {
                        setSharePointSitesError(null);
                      }
                    }}
                  />{" "}
                  Create site on SharePoint
                </span>
                <span className="lookup-hint">
                  {createSharePointSite
                    ? `On save, creates "${clientSharePointSiteDisplayName(form.legalName || "Client Name")}", year/month folders (e.g. 2026/01.Jan26…12.Dec26), and fills Portal Link.`
                    : "Leave unchecked to pick an existing SharePoint site below."}
                </span>
              </label>
            ) : null}

            {!editingId && createSharePointSite ? (
              <label className="field">
                <span className="field-label">SharePoint members</span>
                <textarea
                  className="field-input"
                  rows={3}
                  value={sharePointMemberEmails}
                  disabled={saving}
                  placeholder="email1@fireflyaccountancy.com, email2@fireflyaccountancy.com"
                  onChange={(event) =>
                    setSharePointMemberEmails(event.target.value)
                  }
                />
                <span className="lookup-hint">
                  Optional. Comma-separated Fire Fly emails added as site
                  members. You (signed-in user) are already the owner.
                </span>
              </label>
            ) : null}

            <div className="field">
              <span className="field-label">Portal Link</span>
              {!editingId && createSharePointSite ? (
                <input
                  className="field-input field-input--readonly"
                  type="text"
                  readOnly
                  value={
                    form.portalLink.trim() ||
                    `Will create: ${clientSharePointSiteDisplayName(form.legalName || "Client Name")}`
                  }
                />
              ) : (
                <>
                  <SearchableSelect
                    label="Portal Link"
                    hideLabel
                    value={form.portalLink}
                    options={portalLinkOptions}
                    placeholder={
                      loadingSharePointSites
                        ? "Loading SharePoint sites…"
                        : editingId
                          ? "Select the existing SharePoint site…"
                          : "Select an existing site…"
                    }
                    allowCreate
                    disabled={saving || loadingSharePointSites}
                    onChange={(value) => updateField("portalLink", value)}
                  />
                  {editingId ? (
                    <span className="lookup-hint">
                      Pick the client’s existing SharePoint site (search by
                      name), then Save.
                    </span>
                  ) : null}
                  {sharePointSitesError ? (
                    <p className="form-message error">{sharePointSitesError}</p>
                  ) : null}
                </>
              )}
              {form.portalLink.trim() ? (
                <a
                  className="table-link client-profile-portal-link"
                  href={toWebsiteHref(form.portalLink.trim())}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Open portal link
                </a>
              ) : null}
            </div>
          </FormSection>

          <FormSection
            title="Payroll"
            description="Payroll provider and whether reports are needed from it."
          >
            <SearchableSelect
              label="Payroll Provider"
              value={form.payrollProvider}
              options={payrollProviderOptions}
              allowCreate
              placeholder="Search/Select from the list"
              onChange={(value) => updateField("payrollProvider", value)}
              onCreateOption={(label) => {
                void createLookupValueFromProfile(
                  "payrollProvider",
                  "payrollProvider",
                  "Payroll Provider",
                  label,
                );
              }}
            />

            <fieldset className="field field--radio">
              <legend className="field-label">
                Need reports from payroll provider?
              </legend>
              <div className="radio-row">
                {PAYROLL_REPORTS_NEEDED_OPTIONS.map((option) => (
                  <label key={option} className="radio-option">
                    <input
                      type="radio"
                      name="payrollReportsNeeded"
                      value={option}
                      checked={form.payrollReportsNeeded === option}
                      onChange={() =>
                        updateField(
                          "payrollReportsNeeded",
                          option as PayrollReportsNeeded,
                        )
                      }
                    />
                    {option}
                  </label>
                ))}
              </div>
            </fieldset>

            <label className="field">
              <span className="field-label">Payroll Account Numbers</span>
              <input
                className="field-input"
                value={form.payrollAccountNumbers}
                onChange={(e) =>
                  updateField("payrollAccountNumbers", e.target.value)
                }
              />
            </label>
          </FormSection>

          <FormSection
            title="People"
            description="Owner, day-to-day POC, document-request contacts, and who receives financials."
          >
            <label className="field">
              <span className="field-label">Owner Name</span>
              <input
                className="field-input"
                value={form.ownerName}
                onChange={(e) => updateField("ownerName", e.target.value)}
              />
            </label>

            <label className="field">
              <span className="field-label">Owner Email</span>
              <input
                className="field-input"
                type="text"
                inputMode="email"
                value={form.ownerEmail}
                placeholder="email@example.com, email2@example.com"
                onChange={(e) => updateField("ownerEmail", e.target.value)}
              />
            </label>

            <label className="field">
              <span className="field-label">Owner Phone Number</span>
              <input
                className="field-input"
                value={form.ownerPhone}
                onChange={(e) => updateField("ownerPhone", e.target.value)}
              />
            </label>

            <label className="field">
              <span className="field-label">Point Of Contact</span>
              <input
                className="field-input"
                value={form.pointOfContact}
                onChange={(e) => updateField("pointOfContact", e.target.value)}
              />
            </label>

            <label className="field">
              <span className="field-label">POC Email</span>
              <input
                className="field-input"
                type="text"
                inputMode="email"
                value={form.pocEmail}
                placeholder="email@example.com, email2@example.com"
                onChange={(e) => updateField("pocEmail", e.target.value)}
              />
            </label>

            <label className="field">
              <span className="field-label">POC Phone Number</span>
              <input
                className="field-input"
                value={form.pocPhone}
                onChange={(e) => updateField("pocPhone", e.target.value)}
              />
            </label>

            <p className="field field--wide client-form-subsection-label">
              Document request contacts
            </p>

            <label className="field">
              <span className="field-label">Document Request Contact</span>
              <input
                className="field-input"
                value={form.documentRequestContact}
                onChange={(e) =>
                  updateField("documentRequestContact", e.target.value)
                }
              />
            </label>

            <label className="field">
              <span className="field-label">Document Request Email</span>
              <input
                className="field-input"
                type="text"
                inputMode="email"
                value={form.documentRequestEmail}
                placeholder="email@example.com, email2@example.com"
                onChange={(e) =>
                  updateField("documentRequestEmail", e.target.value)
                }
              />
            </label>

            <label className="field">
              <span className="field-label">Document Request Phone</span>
              <input
                className="field-input"
                value={form.documentRequestPhone}
                onChange={(e) =>
                  updateField("documentRequestPhone", e.target.value)
                }
              />
            </label>

            <p className="field field--wide client-form-subsection-label">
              Contacts who receive financials
            </p>

            <label className="field">
              <span className="field-label">Financials Contact</span>
              <input
                className="field-input"
                value={form.financialsContact}
                onChange={(e) => updateField("financialsContact", e.target.value)}
              />
            </label>

            <label className="field">
              <span className="field-label">Financials Email</span>
              <input
                className="field-input"
                type="text"
                inputMode="email"
                value={form.financialsEmail}
                placeholder="email@example.com, email2@example.com"
                onChange={(e) => updateField("financialsEmail", e.target.value)}
              />
            </label>

            <label className="field">
              <span className="field-label">Financials Phone</span>
              <input
                className="field-input"
                value={form.financialsPhone}
                onChange={(e) => updateField("financialsPhone", e.target.value)}
              />
            </label>
          </FormSection>
          </div>

                  <div className="dialog-actions">
                    <button
                      type="button"
                      className="button secondary"
                      onClick={closeAddForm}
                      disabled={saving}
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      className="button primary"
                      disabled={saving}
                    >
                      {saving ? "Saving…" : "Save"}
                    </button>
                  </div>
                </form>
              </div>
            </div>,
            document.body,
          )
        : null}

      <div className="table-wrap table-wrap--paged" ref={tableWrapRef}>
        {!ready ? (
          <p className="table-empty">Loading…</p>
        ) : profiles.length === 0 ? (
          <p className="table-empty">
            No client profiles yet. Click Add to create one.
          </p>
        ) : visibleProfiles.length === 0 ? (
          <p className="table-empty">No profiles match the current filters.</p>
        ) : (
          <table className="projects-table projects-table--wide projects-table--sticky-cols projects-table--with-check">
            <thead>
              <tr>
                <th className="col-sticky col-sticky-check col-check">
                  <input
                    type="checkbox"
                    checked={
                      pageProfiles.length > 0 &&
                      pageSelectedCount === pageProfiles.length
                    }
                    onChange={toggleSelectAll}
                    aria-label="Select all on this page"
                  />
                </th>
                {TABLE_COLUMNS.map((column, index) => (
                  <th
                    key={column.key}
                    className={
                      index === 0
                        ? "col-sticky col-sticky-1"
                        : index === 1
                          ? "col-sticky col-sticky-2"
                          : undefined
                    }
                  >
                    {column.label}
                  </th>
                ))}
                <th className="col-actions" aria-label="Actions" />
              </tr>
            </thead>
            <tbody>
              {pageProfiles.map((profile) => (
                <tr key={profile.id}>
                  <td className="col-sticky col-sticky-check col-check">
                    <input
                      type="checkbox"
                      checked={selectedIds.includes(profile.id)}
                      onChange={() => toggleSelected(profile.id)}
                      aria-label={`Select ${profile.legalName || profile.uniqueId}`}
                    />
                  </td>
                  {TABLE_COLUMNS.map((column, index) => (
                    <td
                      key={column.key}
                      className={
                        index === 0
                          ? "col-sticky col-sticky-1"
                          : index === 1
                            ? "col-sticky col-sticky-2"
                            : undefined
                      }
                    >
                      {renderProfileCell(
                        column.key,
                        profile[column.key],
                        profile,
                        {
                          onView: (row) => {
                            void openActiveProposal(row, "view");
                          },
                        },
                      )}
                    </td>
                  ))}
                  <td className="col-actions">
                    {canEdit ? (
                      <RowActionsMenu onEdit={() => openEditForm(profile)} />
                    ) : null}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {ready && visibleProfiles.length > 0 ? (
        <div className="pagination-bar">
          <button
            type="button"
            className="button secondary-solid pagination-nav"
            onClick={() => setPage((current) => Math.max(1, current - 1))}
            disabled={currentPage <= 1}
          >
            Prev
          </button>
          <p className="pagination-meta">
            Showing {rangeStart}–{rangeEnd} of {visibleProfiles.length}
          </p>
          <button
            type="button"
            className="button secondary-solid pagination-nav"
            onClick={() =>
              setPage((current) => Math.min(totalPages, current + 1))
            }
            disabled={currentPage >= totalPages}
          >
            Next
          </button>
        </div>
      ) : null}

      <IgnitionProposalViewer
        target={proposalViewer}
        onClose={() => setProposalViewer(null)}
      />
    </section>
  );
}
