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
  DataListToolbar,
  downloadCsv,
  parseSpreadsheetFile,
  SPREADSHEET_IMPORT_ACCEPT,
  uniqueSortedValues,
} from "@/components/DataListToolbar";
import { EngagementTasksDialog } from "@/components/EngagementTasksDialog";
import { RowActionsMenu } from "@/components/RowActionsMenu";
import { SearchableChecklist } from "@/components/SearchableChecklist";
import {
  SearchableSelect,
  type SearchableOption,
} from "@/components/SearchableSelect";
import {
  annualTasksStatusLabel,
  createManualAnnualTask,
  formatAnnualDueDateLabel,
  isManualAnnualTask,
  isProtectedAnnualTask,
  mergeProtectedAnnualTasks,
  normalizeAnnualTasks,
  removeAnnualTask,
  toggleAnnualTaskCompletion,
  updateAnnualTaskFields,
  type AnnualTaskItem,
} from "@/lib/annual-tasks";
import {
  normalizeClientProfile,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import {
  createManualMonthlyTask,
  isManualMonthlyTask,
  isProtectedMonthlyTask,
  mergeProtectedMonthlyTasks,
  monthlyTasksStatusLabel,
  normalizeMonthlyTasks,
  removeMonthlyTask,
  toggleMonthlyTaskCompletion,
  updateMonthlyTaskLabel,
  type MonthlyTaskItem,
} from "@/lib/monthly-tasks";
import {
  createManualOneTimeTask,
  isManualOneTimeTask,
  isProtectedOneTimeTask,
  mergeProtectedOneTimeTasks,
  normalizeOneTimeTasks,
  oneTimeTasksStatusLabel,
  removeOneTimeTask,
  toggleOneTimeTaskCompletion,
  updateOneTimeTaskLabel,
  type OneTimeTaskItem,
} from "@/lib/one-time-tasks";
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
  getDefaultStatus,
  getLookupOptions,
  type LookupDataStore,
  type LookupFieldId,
} from "@/lib/lookup-data";
import { portalFetch } from "@/lib/portal-api";
import type { PortalUser } from "@/lib/portal-users";
import {
  bookkeepingCategoryForCountry,
  createEmptyResourceClientForm,
  formatMultiValue,
  normalizeResourceClient,
  resolveMeetingFrequency,
  type ResourceClientEntry,
  type ResourceClientInput,
} from "@/lib/resource-clients";
import {
  assigneeOptionsFromLookup,
  coerceStaffToEmail,
  formatStaffAssignmentLabel,
  formatStaffAssignmentLabels,
} from "@/lib/staff-assign";
import {
  filterYearlyComplianceOptionsForCountry,
  isCaYearlyComplianceOption,
  isUsYearlyComplianceOption,
  resolveYearlyComplianceTimelines,
} from "@/lib/yearly-compliance";

const RESOURCE_SORT_OPTIONS = [
  { id: "nameAsc", label: "Alphabetically — Client Name" },
  { id: "statusAsc", label: "Status" },
  { id: "industryAsc", label: "Industry" },
  { id: "newest", label: "Newest first" },
  { id: "oldest", label: "Oldest first" },
];

const RESOURCE_EXPORT_COLUMNS: {
  key: keyof ResourceClientEntry;
  label: string;
  format?: (entry: ResourceClientEntry) => string;
}[] = [
  { key: "bookkeepingCategory", label: "Bookkeeping Category" },
  { key: "clientName", label: "Client Name" },
  { key: "industry", label: "Industry" },
  { key: "status", label: "Status" },
  { key: "supervisor", label: "Supervisor" },
  {
    key: "bookkeeperAccountant",
    label: "Bookkeeper/Accountant",
    format: (entry) => formatMultiValue(entry.bookkeeperAccountant),
  },
  {
    key: "bookkeeping",
    label: "Bookkeeping",
    format: (entry) => formatMultiValue(entry.bookkeeping),
  },
  { key: "payroll", label: "Payroll" },
  {
    key: "payrollAndSalesTaxes",
    label: "Payroll and Sales Taxes",
    format: (entry) => formatMultiValue(entry.payrollAndSalesTaxes),
  },
  {
    key: "deliverables",
    label: "Deliverables",
    format: (entry) => formatMultiValue(entry.deliverables),
  },
  {
    key: "yearlyCompliance",
    label: "Yearly Compliance",
    format: (entry) => formatMultiValue(entry.yearlyCompliance),
  },
  { key: "meetingRequirements", label: "Meeting Requirements" },
  { key: "meetingFrequency", label: "Meeting Frequency" },
  { key: "bookkeepingDeadline", label: "Bookkeeping Deadline" },
  { key: "payrollDeadline", label: "Payroll Deadline" },
  { key: "payrollTaxesDeadlines", label: "Payroll Taxes Deadlines" },
  { key: "deliverablesTimelines", label: "Deliverables Timelines" },
  { key: "yearlyComplianceTimelines", label: "Yearly Compliance Timelines" },
];

async function persistClients(clients: ResourceClientEntry[]) {
  const response = await portalFetch("/api/resource-clients", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ clients }),
  });

  if (!response.ok) {
    const data = (await response.json()) as { error?: string };
    throw new Error(data.error || "Failed to save clients.");
  }
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

function resourceSnapshot(entry: ResourceClientEntry) {
  return snapshotFromLabeledFields([
    { label: "Bookkeeping Category", value: entry.bookkeepingCategory },
    { label: "Client Name", value: entry.clientName },
    { label: "Industry", value: entry.industry },
    { label: "Status", value: entry.status },
    { label: "Supervisor", value: entry.supervisor },
    {
      label: "Bookkeeper/Accountant",
      value: formatMultiValue(entry.bookkeeperAccountant),
    },
    { label: "Bookkeeping", value: formatMultiValue(entry.bookkeeping) },
    { label: "Payroll", value: entry.payroll },
    {
      label: "Payroll and Sales Taxes",
      value: formatMultiValue(entry.payrollAndSalesTaxes),
    },
    { label: "Deliverables", value: formatMultiValue(entry.deliverables) },
    {
      label: "Yearly Compliance",
      value: formatMultiValue(entry.yearlyCompliance),
    },
    { label: "Meeting Requirements", value: entry.meetingRequirements },
    { label: "Meeting Frequency", value: entry.meetingFrequency },
    { label: "Bookkeeping Deadline", value: entry.bookkeepingDeadline },
    { label: "Payroll Deadline", value: entry.payrollDeadline },
    {
      label: "Payroll Taxes Deadlines",
      value: entry.payrollTaxesDeadlines,
    },
    { label: "Deliverables Timelines", value: entry.deliverablesTimelines },
    {
      label: "Yearly Compliance Timelines",
      value: entry.yearlyComplianceTimelines,
    },
  ]);
}

export function ResourceAllocationPanel() {
  const { accounts } = useMsal();
  const { canEdit } = useAccess();
  const actor = getActorFromAccount(accounts[0]);
  const [clients, setClients] = useState<ResourceClientEntry[]>([]);
  const [profiles, setProfiles] = useState<ClientProfileEntry[]>([]);
  const [portalUsers, setPortalUsers] = useState<PortalUser[]>([]);
  const [lookups, setLookups] = useState<LookupDataStore | null>(null);
  const [formOpen, setFormOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [form, setForm] = useState<ResourceClientInput>(
    createEmptyResourceClientForm(),
  );
  const [ready, setReady] = useState(false);
  const [mounted, setMounted] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [deleting, setDeleting] = useState(false);
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState("nameAsc");
  const [activeFilters, setActiveFilters] = useState<Record<string, string[]>>(
    {},
  );
  const [importing, setImporting] = useState(false);
  const importInputRef = useRef<HTMLInputElement>(null);
  type TasksDialogKind = "oneTime" | "annual" | "monthly";
  const [tasksDialog, setTasksDialog] = useState<{
    kind: TasksDialogKind;
    clientId: string;
    profileId: string;
  } | null>(null);
  const [oneTimeDraft, setOneTimeDraft] = useState<OneTimeTaskItem[]>([]);
  const [annualDraft, setAnnualDraft] = useState<AnnualTaskItem[]>([]);
  const [monthlyDraft, setMonthlyDraft] = useState<MonthlyTaskItem[]>([]);
  const [tasksNewLabel, setTasksNewLabel] = useState("");
  const [tasksNewDueDate, setTasksNewDueDate] = useState("");
  const [tasksEditingId, setTasksEditingId] = useState<string | null>(null);
  const [tasksEditLabel, setTasksEditLabel] = useState("");
  const [tasksEditDueDate, setTasksEditDueDate] = useState("");
  const [tasksSaving, setTasksSaving] = useState(false);
  const [tasksError, setTasksError] = useState<string | null>(null);

  function profileForClientName(clientName: string): ClientProfileEntry | undefined {
    const needle = clientName.trim().toLowerCase();
    if (!needle) {
      return undefined;
    }
    return profiles.find(
      (entry) => entry.legalName.trim().toLowerCase() === needle,
    );
  }

  function viewListLabel(
    tasks: ReadonlyArray<{ completed: boolean }>,
    status: string,
  ): string {
    if (tasks.length === 0) {
      return "View list";
    }
    return status ? `View list · ${status}` : "View list";
  }

  function closeTasksDialog() {
    if (tasksSaving) {
      return;
    }
    setTasksDialog(null);
    setOneTimeDraft([]);
    setAnnualDraft([]);
    setMonthlyDraft([]);
    setTasksNewLabel("");
    setTasksNewDueDate("");
    setTasksEditingId(null);
    setTasksEditLabel("");
    setTasksEditDueDate("");
    setTasksError(null);
  }

  async function openTasksDialog(
    kind: TasksDialogKind,
    client: ResourceClientEntry,
  ) {
    const profile = profileForClientName(client.clientName);
    if (!profile) {
      setError(
        "Link this client to a Client Profile (matching Legal Name) to manage tasks.",
      );
      return;
    }
    setTasksError(null);
    setTasksNewLabel("");
    setTasksNewDueDate("");
    setTasksEditingId(null);
    setTasksEditLabel("");
    setTasksEditDueDate("");
    setTasksDialog({
      kind,
      clientId: client.id,
      profileId: profile.id,
    });

    const oneTime = normalizeOneTimeTasks(profile.oneTimeTasks);
    const annual = normalizeAnnualTasks(profile.annualTasks);
    const monthly = normalizeMonthlyTasks(profile.monthlyTasks);
    setOneTimeDraft(oneTime);
    setAnnualDraft(annual);
    setMonthlyDraft(monthly);

    const slug = profile.activeProposalSlug.trim();
    if (!canEdit || !slug) {
      return;
    }

    setTasksSaving(true);
    try {
      const endpoint =
        kind === "oneTime"
          ? "/api/client-profiles/one-time-tasks"
          : kind === "annual"
            ? "/api/client-profiles/annual-tasks"
            : "/api/client-profiles/monthly-tasks";
      const body =
        kind === "oneTime"
          ? {
              action: "seed",
              profileId: profile.id,
              forceBookkeeping: client.bookkeeping.some(
                (item) => String(item ?? "").trim().length > 0,
              ),
              bookkeepingServices: client.bookkeeping,
            }
          : {
              action: "seed",
              profileId: profile.id,
              proposalSlug: slug,
            };
      const response = await portalFetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const data = (await response.json()) as {
        error?: string;
        profile?: ClientProfileEntry;
        oneTimeTasks?: OneTimeTaskItem[];
        annualTasks?: AnnualTaskItem[];
        monthlyTasks?: MonthlyTaskItem[];
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to load tasks.");
      }
      if (data.profile) {
        setProfiles((current) =>
          current.map((item) =>
            item.id === data.profile!.id ? data.profile! : item,
          ),
        );
      }
      if (kind === "oneTime") {
        setOneTimeDraft(
          normalizeOneTimeTasks(data.oneTimeTasks ?? data.profile?.oneTimeTasks),
        );
      } else if (kind === "annual") {
        setAnnualDraft(
          normalizeAnnualTasks(data.annualTasks ?? data.profile?.annualTasks),
        );
      } else {
        setMonthlyDraft(
          normalizeMonthlyTasks(
            data.monthlyTasks ?? data.profile?.monthlyTasks,
          ),
        );
      }
    } catch (seedError) {
      setTasksError(
        seedError instanceof Error
          ? seedError.message
          : "Failed to load tasks.",
      );
    } finally {
      setTasksSaving(false);
    }
  }

  async function persistProfileTasks(
    profileId: string,
    patch: Partial<
      Pick<ClientProfileEntry, "oneTimeTasks" | "annualTasks" | "monthlyTasks">
    >,
  ) {
    const profile = profiles.find((item) => item.id === profileId);
    if (!profile) {
      setTasksError("Client profile not found.");
      return;
    }
    const updated = normalizeClientProfile({
      ...profile,
      ...patch,
      oneTimeTasks:
        patch.oneTimeTasks !== undefined
          ? mergeProtectedOneTimeTasks(
              normalizeOneTimeTasks(profile.oneTimeTasks),
              normalizeOneTimeTasks(patch.oneTimeTasks),
            )
          : profile.oneTimeTasks,
      annualTasks:
        patch.annualTasks !== undefined
          ? mergeProtectedAnnualTasks(
              normalizeAnnualTasks(profile.annualTasks),
              normalizeAnnualTasks(patch.annualTasks),
            )
          : profile.annualTasks,
      monthlyTasks:
        patch.monthlyTasks !== undefined
          ? mergeProtectedMonthlyTasks(
              normalizeMonthlyTasks(profile.monthlyTasks),
              normalizeMonthlyTasks(patch.monthlyTasks),
            )
          : profile.monthlyTasks,
    });
    const nextProfiles = profiles.map((item) =>
      item.id === updated.id ? updated : item,
    );
    setTasksSaving(true);
    setTasksError(null);
    try {
      const response = await portalFetch("/api/client-profiles", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ profiles: nextProfiles }),
      });
      const data = (await response.json()) as {
        error?: string;
        profiles?: ClientProfileEntry[];
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to save tasks.");
      }
      setProfiles(Array.isArray(data.profiles) ? data.profiles : nextProfiles);
      if (patch.oneTimeTasks !== undefined) {
        setOneTimeDraft(updated.oneTimeTasks);
      }
      if (patch.annualTasks !== undefined) {
        setAnnualDraft(updated.annualTasks);
      }
      if (patch.monthlyTasks !== undefined) {
        setMonthlyDraft(updated.monthlyTasks);
      }
    } catch (saveError) {
      setTasksError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save tasks.",
      );
    } finally {
      setTasksSaving(false);
    }
  }

  function renderTasksCell(
    client: ResourceClientEntry,
    kind: TasksDialogKind,
  ): ReactNode {
    const profile = profileForClientName(client.clientName);
    if (!profile) {
      return "—";
    }
    const tasks =
      kind === "oneTime"
        ? normalizeOneTimeTasks(profile.oneTimeTasks)
        : kind === "annual"
          ? normalizeAnnualTasks(profile.annualTasks)
          : normalizeMonthlyTasks(profile.monthlyTasks);
    const status =
      kind === "oneTime"
        ? oneTimeTasksStatusLabel(tasks as OneTimeTaskItem[])
        : kind === "annual"
          ? annualTasksStatusLabel(tasks as AnnualTaskItem[])
          : monthlyTasksStatusLabel(tasks as MonthlyTaskItem[]);
    return (
      <button
        type="button"
        className="table-link-button"
        title={`View ${kind === "oneTime" ? "one-time" : kind} tasks`}
        onClick={() => void openTasksDialog(kind, client)}
      >
        {viewListLabel(tasks, status)}
      </button>
    );
  }

  const MULTI_VALUE_KEYS = new Set<keyof ResourceClientEntry>([
    "bookkeeperAccountant",
    "bookkeeping",
    "payrollAndSalesTaxes",
    "deliverables",
    "yearlyCompliance",
  ]);

  function parseMultiCell(raw: string): string[] {
    return raw
      .split(/\s*\+\s*|;|,/)
      .map((part) => part.trim())
      .filter(Boolean);
  }

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function loadAll() {
      try {
        const [clientsRes, profilesRes, lookupsRes, usersRes] =
          await Promise.all([
            fetch("/api/resource-clients"),
            fetch("/api/client-profiles"),
            fetch("/api/lookup-data", { cache: "no-store" }),
            fetch("/api/portal-users", { cache: "no-store" }),
          ]);

        const clientsData = (await clientsRes.json()) as {
          clients?: ResourceClientEntry[];
          error?: string;
        };
        const profilesData = (await profilesRes.json()) as {
          profiles?: ClientProfileEntry[];
          error?: string;
        };
        const lookupsData = (await lookupsRes.json()) as LookupDataStore & {
          error?: string;
        };
        const usersData = (await usersRes.json()) as {
          users?: PortalUser[];
          error?: string;
        };

        if (!clientsRes.ok) {
          throw new Error(clientsData.error || "Failed to load clients.");
        }
        if (!profilesRes.ok) {
          throw new Error(profilesData.error || "Failed to load profiles.");
        }
        if (!lookupsRes.ok) {
          throw new Error(lookupsData.error || "Failed to load lookup data.");
        }

        if (!cancelled) {
          setClients(clientsData.clients ?? []);
          setProfiles(profilesData.profiles ?? []);
          setPortalUsers(
            usersRes.ok && Array.isArray(usersData.users)
              ? usersData.users
              : [],
          );
          setLookups({ fields: lookupsData.fields ?? [] });
          setError(null);
        }
      } catch (loadError) {
        if (!cancelled) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load resource allocation data.",
          );
        }
      } finally {
        if (!cancelled) {
          setReady(true);
        }
      }
    }

    void loadAll();
    return () => {
      cancelled = true;
    };
  }, []);

  const statusOptions = useMemo(
    () => getLookupOptions(lookups, "status"),
    [lookups],
  );
  const industryOptions = useMemo(
    () => getLookupOptions(lookups, "industry"),
    [lookups],
  );
  const supervisorOptions = useMemo(() => {
    const options = assigneeOptionsFromLookup(
      lookups,
      "primaryAssignee",
      portalUsers,
    );
    // Keep current value visible even if removed from Lookup list.
    if (
      form.supervisor.trim() &&
      !options.some((option) => option.value === form.supervisor.trim())
    ) {
      return [
        ...options,
        {
          value: form.supervisor.trim(),
          label:
            formatStaffAssignmentLabel(form.supervisor, portalUsers) ||
            form.supervisor.trim(),
        },
      ];
    }
    return options;
  }, [lookups, portalUsers, form.supervisor]);

  const bookkeeperOptions = useMemo(() => {
    const options = assigneeOptionsFromLookup(
      lookups,
      "secondaryAssignee",
      portalUsers,
    );
    const extras = form.bookkeeperAccountant
      .map((value) => value.trim())
      .filter(Boolean)
      .filter(
        (value) => !options.some((option) => option.value === value),
      )
      .map((value) => ({
        value,
        label: formatStaffAssignmentLabel(value, portalUsers) || value,
      }));
    return [...options, ...extras];
  }, [lookups, portalUsers, form.bookkeeperAccountant]);
  const bookkeepingOptions = useMemo(
    () => getLookupOptions(lookups, "bookkeeping"),
    [lookups],
  );
  const payrollOptions = useMemo(
    () => getLookupOptions(lookups, "payroll"),
    [lookups],
  );
  const payrollTaxOptions = useMemo(
    () => getLookupOptions(lookups, "payrollAndSalesTaxes"),
    [lookups],
  );
  const deliverablesOptions = useMemo(
    () => getLookupOptions(lookups, "deliverables"),
    [lookups],
  );
  const formClientCountry = useMemo(() => {
    const profile = profiles.find(
      (entry) =>
        entry.legalName.trim().toLowerCase() ===
        form.clientName.trim().toLowerCase(),
    );
    return profile?.country;
  }, [profiles, form.clientName]);

  const yearlyComplianceOptions = useMemo(
    () =>
      filterYearlyComplianceOptionsForCountry(
        getLookupOptions(lookups, "yearlyCompliance"),
        formClientCountry,
      ),
    [lookups, formClientCountry],
  );
  const meetingRequirementOptions = useMemo(
    () => getLookupOptions(lookups, "meetingRequirements"),
    [lookups],
  );
  const meetingFrequencyOptions = useMemo(
    () => getLookupOptions(lookups, "meetingFrequency"),
    [lookups],
  );
  const bookkeepingDeadlineOptions = useMemo(
    () => getLookupOptions(lookups, "bookkeepingDeadlines"),
    [lookups],
  );
  const payrollDeadlineOptions = useMemo(
    () => getLookupOptions(lookups, "payrollDeadlines"),
    [lookups],
  );
  const payrollTaxesDeadlineOptions = useMemo(
    () => getLookupOptions(lookups, "payrollTaxesDeadlines"),
    [lookups],
  );
  const deliverablesTimelineOptions = useMemo(
    () => getLookupOptions(lookups, "deliverablesTimelines"),
    [lookups],
  );

  const clientNameOptions = useMemo(() => {
    const usedNames = new Set(
      clients
        .filter((client) => client.id !== editingId)
        .map((client) => client.clientName.trim().toLowerCase())
        .filter(Boolean),
    );

    const options: SearchableOption[] = profiles
      .filter((profile) => profile.legalName.trim())
      .map((profile) => {
        const name = profile.legalName.trim();
        const uniqueId = profile.uniqueId.trim();
        const operatingName = profile.operatingName.trim();
        return {
          value: name,
          label: name,
          meta: uniqueId || undefined,
          keywords: [uniqueId, operatingName].filter(Boolean).join(" "),
        };
      });

    // Prefer first profile when duplicate legal names exist.
    // Hide names already used in Resource Allocation (except current edit row).
    const seen = new Set<string>();
    return options
      .filter((option) => {
        const key = option.value.toLowerCase();
        if (seen.has(key) || usedNames.has(key)) {
          return false;
        }
        seen.add(key);
        return true;
      })
      .sort((a, b) => a.label.localeCompare(b.label));
  }, [profiles, clients, editingId]);

  const meetingFrequencyState = resolveMeetingFrequency(
    form.meetingRequirements,
    form.meetingFrequency,
  );

  function openNewClientForm() {
    setEditingId(null);
    const defaults = createEmptyResourceClientForm(
      getDefaultStatus(statusOptions),
    );
    setForm(defaults);
    setError(null);
    setFormOpen(true);
  }

  function openEditClient(client: ResourceClientEntry) {
    setEditingId(client.id);
    const profile = profiles.find(
      (entry) =>
        entry.legalName.trim().toLowerCase() ===
        client.clientName.trim().toLowerCase(),
    );
    setForm({
      bookkeepingCategory: profile
        ? bookkeepingCategoryForCountry(profile.country)
        : client.bookkeepingCategory,
      clientName: client.clientName,
      industry: profile?.industry.trim() || client.industry,
      status: client.status,
      supervisor: coerceStaffToEmail(client.supervisor, portalUsers),
      bookkeeperAccountant: client.bookkeeperAccountant.map((entry) =>
        coerceStaffToEmail(entry, portalUsers),
      ),
      bookkeeping: [...client.bookkeeping],
      payroll: client.payroll,
      payrollAndSalesTaxes: [...client.payrollAndSalesTaxes],
      deliverables: [...client.deliverables],
      yearlyCompliance: [...client.yearlyCompliance],
      meetingRequirements: client.meetingRequirements,
      meetingFrequency: client.meetingFrequency,
      bookkeepingDeadline: client.bookkeepingDeadline,
      payrollDeadline: client.payrollDeadline,
      payrollTaxesDeadlines: client.payrollTaxesDeadlines,
      deliverablesTimelines: client.deliverablesTimelines,
      yearlyComplianceTimelines: client.yearlyComplianceTimelines,
    });
    setError(null);
    setFormOpen(true);
  }

  function closeForm() {
    if (saving) {
      return;
    }
    setFormOpen(false);
    setEditingId(null);
    setError(null);
  }

  function updateForm<K extends keyof ResourceClientInput>(
    key: K,
    value: ResourceClientInput[K],
  ) {
    setForm((current) => ({ ...current, [key]: value }));
  }

  function handleClientNameChange(clientName: string) {
    const profile = profiles.find(
      (entry) =>
        entry.legalName.trim().toLowerCase() === clientName.trim().toLowerCase(),
    );
    setForm((current) => {
      const country = profile?.country;
      const nextCompliance = profile
        ? current.yearlyCompliance.filter((item) => {
            const normalized = String(country ?? "")
              .trim()
              .toLowerCase();
            if (
              normalized === "ca" ||
              normalized === "canada" ||
              normalized === "canadian"
            ) {
              return isCaYearlyComplianceOption(item);
            }
            if (
              normalized === "not sure" ||
              normalized === "ns"
            ) {
              return true;
            }
            return isUsYearlyComplianceOption(item);
          })
        : current.yearlyCompliance;

      return {
        ...current,
        clientName,
        bookkeepingCategory: profile
          ? bookkeepingCategoryForCountry(profile.country)
          : current.bookkeepingCategory,
        industry: profile ? profile.industry.trim() : current.industry,
        yearlyCompliance: nextCompliance,
        yearlyComplianceTimelines:
          resolveYearlyComplianceTimelines(nextCompliance),
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
        `Delete ${selectedIds.length} selected client${selectedIds.length === 1 ? "" : "s"}?`,
      )
    ) {
      return;
    }

    const removed = clients.filter((client) => selectedIds.includes(client.id));
    const next = clients.filter((client) => !selectedIds.includes(client.id));
    setDeleting(true);
    setError(null);
    try {
      await persistClients(next);
      setClients(next);
      setSelectedIds([]);
      await Promise.all(
        removed.map((client) =>
          logPortalAction({
            actor,
            project: client.clientName,
            entityType: "resourceAllocation",
            entityLabel: client.clientName,
            action: "deleted",
            summary: `Resource allocation deleted: ${client.clientName}`,
            details: Object.entries(resourceSnapshot(client))
              .filter(([, value]) => value.trim())
              .map(([label, value]) => `${label}: ${value}`)
              .join("\n"),
            comment: "Record removed from Resource Allocation.",
          }),
        ),
      );
    } catch (deleteError) {
      setError(
        deleteError instanceof Error
          ? deleteError.message
          : "Failed to delete clients.",
      );
    } finally {
      setDeleting(false);
    }
  }

  async function addLookupSubItem(fieldId: LookupFieldId, label: string) {
    const trimmed = label.trim();
    if (!trimmed) {
      return "";
    }
    if (!lookups) {
      return trimmed;
    }

    const existing = getLookupOptions(lookups, fieldId).find(
      (option) => option.toLowerCase() === trimmed.toLowerCase(),
    );
    if (existing) {
      return existing;
    }

    const fieldName =
      lookups.fields.find((field) => field.id === fieldId)?.name ?? fieldId;

    const next: LookupDataStore = {
      fields: lookups.fields.map((field) =>
        field.id === fieldId
          ? { ...field, items: [...field.items, createLookupItem(trimmed)] }
          : field,
      ),
    };

    const saved = await persistLookupData(next);
    setLookups({ fields: saved.fields ?? [] });
    await logPortalAction({
      actor,
      project: fieldName,
      entityType: "lookupData",
      entityLabel: `${fieldName} / ${trimmed}`,
      action: "created",
      summary: `Lookup item created: ${trimmed}`,
      details: `Field: ${fieldName}\nItem: ${trimmed}\nSource: Resource Allocation`,
      comment: "Lookup sub-item created from Resource Allocation dropdown.",
    });
    return trimmed;
  }

  async function createSingleLookupValue(
    fieldId: LookupFieldId,
    formKey: keyof ResourceClientInput,
    label: string,
  ) {
    const value = await addLookupSubItem(fieldId, label);
    if (!value) {
      return;
    }
    updateForm(formKey, value as never);
  }

  async function createMultiLookupValue(
    fieldId: LookupFieldId,
    formKey: keyof ResourceClientInput,
    label: string,
    currentValues: string[],
  ) {
    const value = await addLookupSubItem(fieldId, label);
    if (!value) {
      return;
    }
    if (
      currentValues.some((item) => item.toLowerCase() === value.toLowerCase())
    ) {
      return;
    }
    updateForm(formKey, [...currentValues, value] as never);
  }

  function handleMeetingRequirementsChange(value: string) {
    const next = resolveMeetingFrequency(value, form.meetingFrequency);
    setForm((current) => ({
      ...current,
      meetingRequirements: value,
      meetingFrequency: next.frequency,
    }));
  }

  function handleYearlyComplianceChange(values: string[]) {
    const nilOption = values.find(
      (value) => value.trim().toLowerCase() === "nil",
    );
    const nextValues = nilOption ? [nilOption] : values;

    setForm((current) => ({
      ...current,
      yearlyCompliance: nextValues,
      yearlyComplianceTimelines: resolveYearlyComplianceTimelines(nextValues),
    }));
  }

  async function handleSave(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    const clientName = form.clientName.trim();
    if (!clientName) {
      setError("Client Name is required.");
      return;
    }

    const duplicate = clients.some(
      (client) =>
        client.id !== editingId &&
        client.clientName.trim().toLowerCase() === clientName.toLowerCase(),
    );
    if (duplicate) {
      setError(
        `Client Name "${clientName}" already exists. Choose a different client.`,
      );
      return;
    }

    const frequency = resolveMeetingFrequency(
      form.meetingRequirements,
      form.meetingFrequency,
    );

    const existing = editingId
      ? clients.find((client) => client.id === editingId)
      : null;

    const profile = profiles.find(
      (entry) =>
        entry.legalName.trim().toLowerCase() === clientName.toLowerCase(),
    );
    const bookkeepingCategory = profile
      ? bookkeepingCategoryForCountry(profile.country)
      : form.bookkeepingCategory;
    const industry = profile?.industry.trim() || form.industry.trim();

    const entry: ResourceClientEntry = {
      id: existing?.id ?? `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      ...form,
      clientName,
      bookkeepingCategory,
      industry,
      status: form.status.trim(),
      supervisor: coerceStaffToEmail(form.supervisor.trim(), portalUsers),
      bookkeeperAccountant: form.bookkeeperAccountant.map((entry) =>
        coerceStaffToEmail(entry, portalUsers),
      ),
      payroll: form.payroll.trim(),
      meetingRequirements: form.meetingRequirements.trim(),
      meetingFrequency: frequency.frequency.trim(),
      bookkeepingDeadline: form.bookkeepingDeadline.trim(),
      payrollDeadline: form.payrollDeadline.trim(),
      payrollTaxesDeadlines: form.payrollTaxesDeadlines.trim(),
      deliverablesTimelines: form.deliverablesTimelines.trim(),
      yearlyComplianceTimelines: resolveYearlyComplianceTimelines(
        form.yearlyCompliance,
      ),
      createdAt: existing?.createdAt ?? new Date().toISOString(),
    };

    const next = editingId
      ? clients.map((client) => (client.id === editingId ? entry : client))
      : [entry, ...clients];
    setSaving(true);
    setError(null);

    try {
      await persistClients(next);
      setClients(next);
      setFormOpen(false);
      setEditingId(null);
      setForm(createEmptyResourceClientForm(getDefaultStatus(statusOptions)));

      // Auto-seed One-Time / Annual / Monthly from the linked proposal (empty if none).
      if (profile?.id && profile.activeProposalSlug.trim()) {
        void Promise.all([
          portalFetch("/api/client-profiles/one-time-tasks", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              action: "seed",
              profileId: profile.id,
              forceBookkeeping: entry.bookkeeping.some(
                (item) => String(item ?? "").trim().length > 0,
              ),
              bookkeepingServices: entry.bookkeeping,
            }),
          }),
          portalFetch("/api/client-profiles/annual-tasks", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              action: "seed",
              profileId: profile.id,
              proposalSlug: profile.activeProposalSlug,
            }),
          }),
          portalFetch("/api/client-profiles/monthly-tasks", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              action: "seed",
              profileId: profile.id,
              proposalSlug: profile.activeProposalSlug,
            }),
          }),
        ])
          .then(async (responses) => {
            const bodies = await Promise.all(
              responses.map(async (response) => {
                try {
                  return (await response.json()) as {
                    profile?: ClientProfileEntry;
                  };
                } catch {
                  return {};
                }
              }),
            );
            const latest = bodies
              .map((body) => body.profile)
              .filter((item): item is ClientProfileEntry => Boolean(item))
              .at(-1);
            if (latest) {
              setProfiles((current) =>
                current.map((item) =>
                  item.id === latest.id ? latest : item,
                ),
              );
            }
          })
          .catch(() => undefined);
      }

      const after = resourceSnapshot(entry);
      if (existing) {
        const details = buildChangeDetails(resourceSnapshot(existing), after);
        await logPortalAction({
          actor,
          project: entry.clientName,
          entityType: "resourceAllocation",
          entityLabel: entry.clientName,
          action: "updated",
          summary: `Resource allocation updated: ${entry.clientName}`,
          details: details || "No field values changed.",
          comment: "Changes saved on Resource Allocation.",
        });
      } else {
        await logPortalAction({
          actor,
          project: entry.clientName,
          entityType: "resourceAllocation",
          entityLabel: entry.clientName,
          action: "created",
          summary: `Resource allocation created: ${entry.clientName}`,
          details: Object.entries(after)
            .filter(([, value]) => value.trim())
            .map(([label, value]) => `${label}: ${value}`)
            .join("\n"),
          comment: "New resource allocation added.",
        });
      }
    } catch (saveError) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save client.",
      );
    } finally {
      setSaving(false);
    }
  }

  const toolbarFilters = useMemo(
    () => [
      {
        id: "bookkeepingCategory",
        label: "Category",
        options: uniqueSortedValues(
          clients.map((client) => client.bookkeepingCategory),
        ),
      },
      {
        id: "industry",
        label: "Industry",
        options: uniqueSortedValues(clients.map((client) => client.industry)),
      },
      {
        id: "status",
        label: "Status",
        options: uniqueSortedValues(clients.map((client) => client.status)),
      },
      {
        id: "supervisor",
        label: "Supervisor",
        options: uniqueSortedValues(clients.map((client) => client.supervisor)),
      },
      {
        id: "payroll",
        label: "Payroll",
        options: uniqueSortedValues(clients.map((client) => client.payroll)),
      },
    ],
    [clients],
  );

  const visibleClients = useMemo(() => {
    const query = search.trim().toLowerCase();
    let rows = clients.filter((client) => {
      if (query) {
        const haystack = [
          client.clientName,
          client.industry,
          client.status,
          client.supervisor,
          client.bookkeepingCategory,
          client.payroll,
          formatMultiValue(client.bookkeeperAccountant),
          formatMultiValue(client.bookkeeping),
        ]
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
          client[filterId as keyof ResourceClientEntry] ?? "",
        ).trim();
        if (!selected.includes(value)) {
          return false;
        }
      }
      return true;
    });

    rows = [...rows].sort((a, b) => {
      switch (sort) {
        case "statusAsc":
          return (
            a.status.localeCompare(b.status, undefined, {
              sensitivity: "base",
            }) ||
            a.clientName.localeCompare(b.clientName, undefined, {
              sensitivity: "base",
            })
          );
        case "industryAsc":
          return (
            a.industry.localeCompare(b.industry, undefined, {
              sensitivity: "base",
            }) ||
            a.clientName.localeCompare(b.clientName, undefined, {
              sensitivity: "base",
            })
          );
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
          return a.clientName.localeCompare(b.clientName, undefined, {
            sensitivity: "base",
          });
      }
    });

    return rows;
  }, [clients, search, activeFilters, sort]);

  function handleFilterChange(filterId: string, values: string[]) {
    setActiveFilters((current) => ({ ...current, [filterId]: values }));
  }

  function toggleSelectAll() {
    const visibleIds = visibleClients.map((client) => client.id);
    const allVisibleSelected =
      visibleIds.length > 0 &&
      visibleIds.every((id) => selectedIds.includes(id));
    if (allVisibleSelected) {
      setSelectedIds((current) =>
        current.filter((id) => !visibleIds.includes(id)),
      );
      return;
    }
    setSelectedIds((current) => [...new Set([...current, ...visibleIds])]);
  }

  function handleExport() {
    const header = RESOURCE_EXPORT_COLUMNS.map((column) => column.label);
    const rows = visibleClients.map((client) =>
      RESOURCE_EXPORT_COLUMNS.map((column) =>
        column.format
          ? column.format(client)
          : String(client[column.key] ?? ""),
      ),
    );
    downloadCsv("resource-allocation.csv", [header, ...rows]);
  }

  function openImportPicker() {
    importInputRef.current?.click();
  }

  async function handleImportFile(file: File) {
    setImporting(true);
    setError(null);
    try {
      const rows = await parseSpreadsheetFile(file);
      if (rows.length < 2) {
        throw new Error(
          "File must include a header row and at least one data row.",
        );
      }

      const headers = rows[0].map((header) => header.trim().toLowerCase());
      const labelToKey = new Map(
        RESOURCE_EXPORT_COLUMNS.map((column) => [
          column.label.trim().toLowerCase(),
          column.key,
        ]),
      );
      for (const column of RESOURCE_EXPORT_COLUMNS) {
        labelToKey.set(column.key.toLowerCase(), column.key);
      }

      const columnIndexes: {
        key: keyof ResourceClientEntry;
        index: number;
      }[] = [];
      headers.forEach((header, index) => {
        const key = labelToKey.get(header);
        if (key) {
          columnIndexes.push({ key, index });
        }
      });

      if (columnIndexes.length === 0) {
        throw new Error(
          "No recognized columns. Export first and use those headers.",
        );
      }

      const next = [...clients];
      let created = 0;
      let updated = 0;

      for (const row of rows.slice(1)) {
        const values: Partial<ResourceClientEntry> = {};
        for (const { key, index } of columnIndexes) {
          const cell = (row[index] ?? "").trim();
          if (MULTI_VALUE_KEYS.has(key)) {
            (values as Record<string, string[]>)[key] = parseMultiCell(cell);
          } else if (key === "bookkeepingCategory") {
            const cat = cell.toUpperCase();
            if (cat === "ACC1" || cat === "ACC2" || cat === "ACC3") {
              values.bookkeepingCategory = cat;
            }
          } else {
            (values as Record<string, string>)[key] = cell;
          }
        }

        const clientName = String(values.clientName ?? "").trim();
        if (!clientName) {
          continue;
        }

        const existingIndex = next.findIndex(
          (client) =>
            client.clientName.trim().toLowerCase() === clientName.toLowerCase(),
        );

        if (existingIndex >= 0) {
          const existing = next[existingIndex]!;
          next[existingIndex] = normalizeResourceClient({
            ...existing,
            ...values,
            id: existing.id,
            clientName,
            createdAt: existing.createdAt,
          });
          updated += 1;
        } else {
          next.unshift(
            normalizeResourceClient({
              id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
              ...createEmptyResourceClientForm(getDefaultStatus(statusOptions)),
              ...values,
              clientName,
              createdAt: new Date().toISOString(),
            }),
          );
          created += 1;
        }
      }

      if (created === 0 && updated === 0) {
        throw new Error("No valid rows found. Client Name is required.");
      }

      await persistClients(next);
      setClients(next);
      setError(null);
      await logPortalAction({
        actor,
        entityType: "resourceAllocation",
        action: "updated",
        project: "Resource Allocation",
        entityLabel: "Excel/CSV Import",
        summary: `Imported resource clients (${created} added, ${updated} updated)`,
        details: `File: ${file.name}`,
        comment: "Bulk import from Excel/CSV.",
      });
    } catch (importError) {
      setError(
        importError instanceof Error
          ? importError.message
          : "Failed to import resource clients.",
      );
    } finally {
      setImporting(false);
      if (importInputRef.current) {
        importInputRef.current.value = "";
      }
    }
  }

  const visibleSelectedCount = visibleClients.filter((client) =>
    selectedIds.includes(client.id),
  ).length;

  return (
    <section className="content-panel content-panel--actions">
      <DataListToolbar
        search={search}
        onSearchChange={setSearch}
        searchPlaceholder="Client name, industry, status, supervisor..."
        sort={sort}
        onSortChange={setSort}
        sortOptions={RESOURCE_SORT_OPTIONS}
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
              disabled={visibleClients.length === 0}
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
                  accept={SPREADSHEET_IMPORT_ACCEPT}
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
                  onClick={openNewClientForm}
                >
                  New Client
                </button>
              </>
            ) : null}
          </>
        }
      />

      {error && !formOpen ? (
        <p className="form-message error">{error}</p>
      ) : null}

      <div className="table-wrap">
        {!ready ? (
          <p className="table-empty">Loading…</p>
        ) : clients.length === 0 ? (
          <p className="table-empty">
            No clients yet. Click New Client to create one.
          </p>
        ) : visibleClients.length === 0 ? (
          <p className="table-empty">No clients match the current filters.</p>
        ) : (
          <table className="projects-table projects-table--wide projects-table--sticky-cols projects-table--with-check">
            <thead>
              <tr>
                <th className="col-sticky col-sticky-check col-check">
                  <input
                    type="checkbox"
                    checked={
                      visibleClients.length > 0 &&
                      visibleSelectedCount === visibleClients.length
                    }
                    onChange={toggleSelectAll}
                    aria-label="Select all"
                  />
                </th>
                <th className="col-sticky col-sticky-1">Bookkeeping Category</th>
                <th className="col-sticky col-sticky-2">Client Name</th>
                <th>Industry</th>
                <th>Status</th>
                <th>Supervisor</th>
                <th>Bookkeeper/Accountant</th>
                <th>Bookkeeping services</th>
                <th>Payroll</th>
                <th>Payroll and Sales Taxes</th>
                <th>Deliverables</th>
                <th>Yearly Compliance</th>
                <th>Meeting Requirements</th>
                <th>Meeting Frequency</th>
                <th>Bookkeeping Deadline</th>
                <th>Payroll Deadline</th>
                <th>Payroll Taxes Deadlines</th>
                <th>Deliverables Timelines</th>
                <th>Yearly compliance Timelines</th>
                <th>One-Time Tasks</th>
                <th>Annual Tasks</th>
                <th>Monthly Tasks</th>
                <th className="col-actions" aria-label="Actions" />
              </tr>
            </thead>
            <tbody>
              {visibleClients.map((client) => (
                <tr key={client.id}>
                  <td className="col-sticky col-sticky-check col-check">
                    <input
                      type="checkbox"
                      checked={selectedIds.includes(client.id)}
                      onChange={() => toggleSelected(client.id)}
                      aria-label={`Select ${client.clientName}`}
                    />
                  </td>
                  <td className="col-sticky col-sticky-1">
                    {client.bookkeepingCategory}
                  </td>
                  <td className="col-sticky col-sticky-2">{client.clientName}</td>
                  <td>{client.industry || "—"}</td>
                  <td>{client.status || "—"}</td>
                  <td>
                    {formatStaffAssignmentLabel(
                      client.supervisor,
                      portalUsers,
                    ) || "—"}
                  </td>
                  <td>
                    {formatStaffAssignmentLabels(
                      client.bookkeeperAccountant,
                      portalUsers,
                    ) || "—"}
                  </td>
                  <td>{formatMultiValue(client.bookkeeping) || "—"}</td>
                  <td>{client.payroll || "—"}</td>
                  <td>
                    {formatMultiValue(client.payrollAndSalesTaxes) || "—"}
                  </td>
                  <td>{formatMultiValue(client.deliverables) || "—"}</td>
                  <td>{formatMultiValue(client.yearlyCompliance) || "—"}</td>
                  <td>{client.meetingRequirements || "—"}</td>
                  <td>{client.meetingFrequency || "—"}</td>
                  <td>{client.bookkeepingDeadline || "—"}</td>
                  <td>{client.payrollDeadline || "—"}</td>
                  <td>{client.payrollTaxesDeadlines || "—"}</td>
                  <td>{client.deliverablesTimelines || "—"}</td>
                  <td className="cell-multiline">
                    {client.yearlyComplianceTimelines || "—"}
                  </td>
                  <td>{renderTasksCell(client, "oneTime")}</td>
                  <td>{renderTasksCell(client, "annual")}</td>
                  <td>{renderTasksCell(client, "monthly")}</td>
                  <td className="col-actions">
                    {canEdit ? (
                      <RowActionsMenu onEdit={() => openEditClient(client)} />
                    ) : null}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {formOpen && mounted
        ? createPortal(
            <div className="dialog-backdrop" role="presentation">
              <div
                className="dialog-panel dialog-panel--wide"
                role="dialog"
                aria-modal="true"
                aria-labelledby="resource-client-dialog-title"
              >
                <div className="dialog-header">
                  <h3
                    id="resource-client-dialog-title"
                    className="dialog-title"
                  >
                    {editingId ? "Edit Client" : "New Client"}
                  </h3>
                  <button
                    type="button"
                    className="dialog-close"
                    aria-label="Close"
                    onClick={closeForm}
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
                  <div className="client-form-grid">
                    <SearchableSelect
                      label="Client Name"
                      value={form.clientName}
                      options={clientNameOptions}
                      required
                      placeholder="Search by ID or name"
                      onChange={handleClientNameChange}
                    />

                    <label className="field">
                      <span className="field-label">Bookkeeping Category</span>
                      <input
                        className="field-input field-input--readonly"
                        value={
                          form.clientName.trim()
                            ? form.bookkeepingCategory
                            : ""
                        }
                        placeholder="Select a client name"
                        readOnly
                        tabIndex={-1}
                      />
                    </label>

                    <label className="field">
                      <span className="field-label">Industry</span>
                      <input
                        className="field-input field-input--readonly"
                        value={form.clientName.trim() ? form.industry : ""}
                        placeholder="Select a client name"
                        readOnly
                        tabIndex={-1}
                      />
                    </label>

                    <SearchableSelect
                      label="Status"
                      value={form.status}
                      options={statusOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(value) => updateForm("status", value)}
                      onCreateOption={(label) =>
                        createSingleLookupValue("status", "status", label)
                      }
                    />

                    <SearchableSelect
                      label="Supervisor (Primary)"
                      value={form.supervisor}
                      options={supervisorOptions}
                      placeholder="Search Primary Assignee list…"
                      onChange={(value) => updateForm("supervisor", value)}
                    />
                    <span className="lookup-hint field--wide">
                      Options come from Lookup → Primary Assignee (picked from
                      User Management). This mailbox sends client schedule
                      email.
                    </span>

                    <SearchableChecklist
                      label="Bookkeeper/Accountant (Secondary)"
                      values={form.bookkeeperAccountant}
                      options={bookkeeperOptions}
                      placeholder="Search Secondary Assignee list…"
                      onChange={(values) =>
                        updateForm("bookkeeperAccountant", values)
                      }
                    />
                    <span className="lookup-hint field--wide">
                      Options come from Lookup → Secondary Assignee. Secondary
                      can view schedules but cannot send email.
                    </span>

                    <SearchableChecklist
                      label="Bookkeeping services"
                      values={form.bookkeeping}
                      options={bookkeepingOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(values) => updateForm("bookkeeping", values)}
                      onCreateOption={(label) =>
                        createMultiLookupValue(
                          "bookkeeping",
                          "bookkeeping",
                          label,
                          form.bookkeeping,
                        )
                      }
                    />

                    <SearchableSelect
                      label="Payroll"
                      value={form.payroll}
                      options={payrollOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(value) => updateForm("payroll", value)}
                      onCreateOption={(label) =>
                        createSingleLookupValue("payroll", "payroll", label)
                      }
                    />

                    <SearchableChecklist
                      label="Payroll and Sales Taxes"
                      values={form.payrollAndSalesTaxes}
                      options={payrollTaxOptions}
                      allowCreate
                      exclusiveOption="NIL"
                      placeholder="Search/Select from the list"
                      onChange={(values) => {
                        const nilOption = values.find(
                          (value) => value.trim().toLowerCase() === "nil",
                        );
                        updateForm(
                          "payrollAndSalesTaxes",
                          nilOption ? [nilOption] : values,
                        );
                      }}
                      onCreateOption={(label) => {
                        if (
                          form.payrollAndSalesTaxes.some(
                            (value) => value.trim().toLowerCase() === "nil",
                          )
                        ) {
                          return;
                        }
                        return createMultiLookupValue(
                          "payrollAndSalesTaxes",
                          "payrollAndSalesTaxes",
                          label,
                          form.payrollAndSalesTaxes,
                        );
                      }}
                    />

                    <SearchableChecklist
                      label="Deliverables"
                      values={form.deliverables}
                      options={deliverablesOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(values) => updateForm("deliverables", values)}
                      onCreateOption={(label) =>
                        createMultiLookupValue(
                          "deliverables",
                          "deliverables",
                          label,
                          form.deliverables,
                        )
                      }
                    />

                    <SearchableChecklist
                      label="Yearly Compliance"
                      values={form.yearlyCompliance}
                      options={yearlyComplianceOptions}
                      allowCreate
                      exclusiveOption="NIL"
                      placeholder={
                        formClientCountry === "CA"
                          ? "Canadian forms (T1, T2, …) or NIL"
                          : formClientCountry === "US"
                            ? "US forms (1120, 1120S, …) or NIL"
                            : "Search/Select from the list"
                      }
                      onChange={handleYearlyComplianceChange}
                      onCreateOption={async (label) => {
                        if (
                          form.yearlyCompliance.some(
                            (item) => item.trim().toLowerCase() === "nil",
                          )
                        ) {
                          return;
                        }
                        const allowed =
                          formClientCountry === "CA"
                            ? isCaYearlyComplianceOption(label)
                            : formClientCountry === "US" || !formClientCountry
                              ? isUsYearlyComplianceOption(label)
                              : true;
                        if (!allowed) {
                          setError(
                            formClientCountry === "CA"
                              ? "Canadian clients can only use letter-based compliance codes (T1, T2, …) or NIL."
                              : "US clients can only use numeric compliance codes (1120, 1120S, …) or NIL.",
                          );
                          return;
                        }
                        const value = await addLookupSubItem(
                          "yearlyCompliance",
                          label,
                        );
                        if (!value) {
                          return;
                        }
                        if (
                          form.yearlyCompliance.some(
                            (item) =>
                              item.toLowerCase() === value.toLowerCase(),
                          )
                        ) {
                          return;
                        }
                        handleYearlyComplianceChange([
                          ...form.yearlyCompliance,
                          value,
                        ]);
                      }}
                    />

                    <SearchableSelect
                      label="Meeting Requirements"
                      value={form.meetingRequirements}
                      options={meetingRequirementOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={handleMeetingRequirementsChange}
                      onCreateOption={async (label) => {
                        const value = await addLookupSubItem(
                          "meetingRequirements",
                          label,
                        );
                        if (value) {
                          handleMeetingRequirementsChange(value);
                        }
                      }}
                    />

                    <SearchableSelect
                      label="Meeting Frequency"
                      value={
                        meetingFrequencyState.locked
                          ? meetingFrequencyState.frequency
                          : form.meetingFrequency
                      }
                      options={meetingFrequencyOptions}
                      allowCreate={!meetingFrequencyState.locked}
                      disabled={meetingFrequencyState.locked}
                      placeholder={
                        meetingFrequencyState.locked
                          ? "Set by Meeting Requirements"
                          : "Search/Select from the list"
                      }
                      onChange={(value) =>
                        updateForm("meetingFrequency", value)
                      }
                      onCreateOption={(label) =>
                        createSingleLookupValue(
                          "meetingFrequency",
                          "meetingFrequency",
                          label,
                        )
                      }
                    />

                    <SearchableSelect
                      label="Bookkeeping"
                      value={form.bookkeepingDeadline}
                      options={bookkeepingDeadlineOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(value) =>
                        updateForm("bookkeepingDeadline", value)
                      }
                      onCreateOption={(label) =>
                        createSingleLookupValue(
                          "bookkeepingDeadlines",
                          "bookkeepingDeadline",
                          label,
                        )
                      }
                    />

                    <SearchableSelect
                      label="Payroll"
                      value={form.payrollDeadline}
                      options={payrollDeadlineOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(value) => updateForm("payrollDeadline", value)}
                      onCreateOption={(label) =>
                        createSingleLookupValue(
                          "payrollDeadlines",
                          "payrollDeadline",
                          label,
                        )
                      }
                    />

                    <SearchableSelect
                      label="Payroll Taxes Deadlines"
                      value={form.payrollTaxesDeadlines}
                      options={payrollTaxesDeadlineOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(value) =>
                        updateForm("payrollTaxesDeadlines", value)
                      }
                      onCreateOption={(label) =>
                        createSingleLookupValue(
                          "payrollTaxesDeadlines",
                          "payrollTaxesDeadlines",
                          label,
                        )
                      }
                    />

                    <SearchableSelect
                      label="Deliverables Timelines"
                      value={form.deliverablesTimelines}
                      options={deliverablesTimelineOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(value) =>
                        updateForm("deliverablesTimelines", value)
                      }
                      onCreateOption={(label) =>
                        createSingleLookupValue(
                          "deliverablesTimelines",
                          "deliverablesTimelines",
                          label,
                        )
                      }
                    />

                    <label className="field field--wide">
                      <span className="field-label">
                        Yearly compliance Timelines
                      </span>
                      <textarea
                        className="field-input field-input--readonly field-textarea-readonly"
                        value={form.yearlyComplianceTimelines}
                        readOnly
                        rows={Math.max(
                          2,
                          form.yearlyComplianceTimelines.split("\n").length,
                        )}
                        placeholder="Auto-filled from Yearly Compliance"
                      />
                    </label>
                  </div>

                  <div className="dialog-actions">
                    <button
                      type="button"
                      className="button secondary"
                      onClick={closeForm}
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

      {tasksDialog?.kind === "oneTime"
        ? createPortal(
            <EngagementTasksDialog
              title="One-Time Tasks"
              titleId="ra-one-time-tasks-title"
              statusLabel={oneTimeTasksStatusLabel(oneTimeDraft) || "Pending"}
              itemCount={oneTimeDraft.length}
              saving={tasksSaving}
              error={tasksError}
              canEdit={canEdit}
              onClose={closeTasksDialog}
              intro={
                <>
                  From the linked Ignition proposal when present; empty if no
                  proposal. Defaults cannot be edited or deleted.
                </>
              }
              emptyHint={
                <>
                  <p>No one-time tasks yet.</p>
                  <p>
                    QBO Setup fills in when the proposal includes bookkeeping
                    (or Resource Allocation Bookkeeping is set).
                  </p>
                </>
              }
              addRow={
                <div className="one-time-tasks-add">
                  <label className="field one-time-tasks-add-field">
                    <span className="field-label">Add another task</span>
                    <input
                      className="field-input"
                      value={tasksNewLabel}
                      disabled={tasksSaving}
                      placeholder="Type a new one-time task…"
                      onChange={(event) => setTasksNewLabel(event.target.value)}
                    />
                  </label>
                  <button
                    type="button"
                    className="button secondary-solid"
                    disabled={tasksSaving || !tasksNewLabel.trim()}
                    onClick={() => {
                      const label = tasksNewLabel.trim();
                      if (!label || !tasksDialog) {
                        return;
                      }
                      setTasksNewLabel("");
                      void persistProfileTasks(tasksDialog.profileId, {
                        oneTimeTasks: [
                          ...oneTimeDraft,
                          createManualOneTimeTask(label),
                        ],
                      });
                    }}
                  >
                    Add
                  </button>
                </div>
              }
            >
              {oneTimeDraft.map((task) => {
                const manual = isManualOneTimeTask(task);
                const editing = tasksEditingId === task.id;
                return (
                  <li
                    key={task.id}
                    className={
                      task.completed
                        ? "one-time-tasks-item one-time-tasks-item--done"
                        : "one-time-tasks-item"
                    }
                  >
                    <div className="one-time-tasks-item-row">
                      <label className="one-time-tasks-check">
                        <input
                          type="checkbox"
                          checked={task.completed}
                          disabled={!canEdit || tasksSaving}
                          onChange={(event) => {
                            if (!tasksDialog) {
                              return;
                            }
                            void persistProfileTasks(tasksDialog.profileId, {
                              oneTimeTasks: toggleOneTimeTaskCompletion(
                                oneTimeDraft,
                                task.id,
                                event.target.checked,
                                actor,
                              ),
                            });
                          }}
                        />
                        {editing ? (
                          <input
                            className="field-input one-time-tasks-edit-input"
                            value={tasksEditLabel}
                            disabled={tasksSaving}
                            autoFocus
                            onChange={(event) =>
                              setTasksEditLabel(event.target.value)
                            }
                          />
                        ) : (
                          <span className="one-time-tasks-label">
                            {task.label}
                            {isProtectedOneTimeTask(task) ? (
                              <span className="one-time-tasks-lock">
                                {" "}
                                · Default
                              </span>
                            ) : null}
                          </span>
                        )}
                      </label>
                      {canEdit && manual ? (
                        <div className="one-time-tasks-item-actions">
                          {editing ? (
                            <>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={
                                  tasksSaving || !tasksEditLabel.trim()
                                }
                                onClick={() => {
                                  if (!tasksDialog) {
                                    return;
                                  }
                                  const next = updateOneTimeTaskLabel(
                                    oneTimeDraft,
                                    task.id,
                                    tasksEditLabel,
                                  );
                                  setTasksEditingId(null);
                                  setTasksEditLabel("");
                                  void persistProfileTasks(
                                    tasksDialog.profileId,
                                    { oneTimeTasks: next },
                                  );
                                }}
                              >
                                Save
                              </button>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={tasksSaving}
                                onClick={() => {
                                  setTasksEditingId(null);
                                  setTasksEditLabel("");
                                }}
                              >
                                Cancel
                              </button>
                            </>
                          ) : (
                            <>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={tasksSaving}
                                onClick={() => {
                                  setTasksEditingId(task.id);
                                  setTasksEditLabel(task.label);
                                }}
                              >
                                Edit
                              </button>
                              <button
                                type="button"
                                className="table-link-button one-time-tasks-delete"
                                disabled={tasksSaving}
                                onClick={() => {
                                  if (!tasksDialog) {
                                    return;
                                  }
                                  void persistProfileTasks(
                                    tasksDialog.profileId,
                                    {
                                      oneTimeTasks: removeOneTimeTask(
                                        oneTimeDraft,
                                        task.id,
                                      ),
                                    },
                                  );
                                }}
                              >
                                Delete
                              </button>
                            </>
                          )}
                        </div>
                      ) : null}
                    </div>
                    {task.completed ? (
                      <p className="one-time-tasks-meta">
                        Completed{" "}
                        {task.completedAt
                          ? new Date(task.completedAt).toLocaleString()
                          : ""}
                        {task.completedByName
                          ? ` · ${task.completedByName}`
                          : ""}
                      </p>
                    ) : null}
                  </li>
                );
              })}
            </EngagementTasksDialog>,
            document.body,
          )
        : null}

      {tasksDialog?.kind === "annual"
        ? createPortal(
            <EngagementTasksDialog
              title="Annual Tasks"
              titleId="ra-annual-tasks-title"
              statusLabel={annualTasksStatusLabel(annualDraft) || "Pending"}
              itemCount={annualDraft.length}
              saving={tasksSaving}
              error={tasksError}
              canEdit={canEdit}
              onClose={closeTasksDialog}
              intro={
                <>
                  Tax / Tax Compliance services from the proposal, with due
                  dates. Empty when no proposal. Defaults cannot be edited or
                  deleted.
                </>
              }
              emptyHint={
                <>
                  <p>No annual tasks yet.</p>
                  <p>
                    Items appear when the linked proposal includes Tax services
                    or Tax Compliance Services.
                  </p>
                </>
              }
              addRow={
                <div className="one-time-tasks-add annual-tasks-add">
                  <label className="field one-time-tasks-add-field">
                    <span className="field-label">Add another task</span>
                    <input
                      className="field-input"
                      value={tasksNewLabel}
                      disabled={tasksSaving}
                      placeholder="Type a new annual task…"
                      onChange={(event) => setTasksNewLabel(event.target.value)}
                    />
                  </label>
                  <label className="field annual-tasks-add-due-field">
                    <span className="field-label">Due date</span>
                    <input
                      type="date"
                      className="field-input"
                      value={tasksNewDueDate}
                      disabled={tasksSaving}
                      onChange={(event) =>
                        setTasksNewDueDate(event.target.value)
                      }
                    />
                  </label>
                  <button
                    type="button"
                    className="button secondary-solid"
                    disabled={tasksSaving || !tasksNewLabel.trim()}
                    onClick={() => {
                      const label = tasksNewLabel.trim();
                      if (!label || !tasksDialog) {
                        return;
                      }
                      setTasksNewLabel("");
                      const due = tasksNewDueDate;
                      setTasksNewDueDate("");
                      void persistProfileTasks(tasksDialog.profileId, {
                        annualTasks: [
                          ...annualDraft,
                          createManualAnnualTask(label, due),
                        ],
                      });
                    }}
                  >
                    Add
                  </button>
                </div>
              }
            >
              <li className="annual-tasks-header-row" aria-hidden="true">
                <span className="annual-tasks-col-task">Task</span>
                <span className="annual-tasks-col-due">Due date</span>
                <span className="annual-tasks-col-actions" />
              </li>
              {annualDraft.map((task) => {
                const manual = isManualAnnualTask(task);
                const editing = tasksEditingId === task.id;
                return (
                  <li
                    key={task.id}
                    className={
                      task.completed
                        ? "one-time-tasks-item one-time-tasks-item--done"
                        : "one-time-tasks-item"
                    }
                  >
                    <div className="one-time-tasks-item-row annual-tasks-item-row">
                      <label className="one-time-tasks-check annual-tasks-col-task">
                        <input
                          type="checkbox"
                          checked={task.completed}
                          disabled={!canEdit || tasksSaving}
                          onChange={(event) => {
                            if (!tasksDialog) {
                              return;
                            }
                            void persistProfileTasks(tasksDialog.profileId, {
                              annualTasks: toggleAnnualTaskCompletion(
                                annualDraft,
                                task.id,
                                event.target.checked,
                                actor,
                              ),
                            });
                          }}
                        />
                        {editing ? (
                          <input
                            className="field-input one-time-tasks-edit-input"
                            value={tasksEditLabel}
                            disabled={tasksSaving}
                            autoFocus
                            onChange={(event) =>
                              setTasksEditLabel(event.target.value)
                            }
                          />
                        ) : (
                          <span className="one-time-tasks-label">
                            {task.label}
                            {isProtectedAnnualTask(task) ? (
                              <span className="one-time-tasks-lock">
                                {" "}
                                · Default
                              </span>
                            ) : null}
                          </span>
                        )}
                      </label>
                      <div className="annual-tasks-col-due">
                        {editing ? (
                          <input
                            type="date"
                            className="field-input annual-tasks-due-input"
                            value={tasksEditDueDate}
                            disabled={tasksSaving}
                            onChange={(event) =>
                              setTasksEditDueDate(event.target.value)
                            }
                          />
                        ) : (
                          <span className="annual-tasks-due-label">
                            {formatAnnualDueDateLabel(task.dueDate)}
                          </span>
                        )}
                      </div>
                      {canEdit && manual ? (
                        <div className="one-time-tasks-item-actions annual-tasks-col-actions">
                          {editing ? (
                            <>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={
                                  tasksSaving || !tasksEditLabel.trim()
                                }
                                onClick={() => {
                                  if (!tasksDialog) {
                                    return;
                                  }
                                  const next = updateAnnualTaskFields(
                                    annualDraft,
                                    task.id,
                                    {
                                      label: tasksEditLabel,
                                      dueDate: tasksEditDueDate,
                                    },
                                  );
                                  setTasksEditingId(null);
                                  setTasksEditLabel("");
                                  setTasksEditDueDate("");
                                  void persistProfileTasks(
                                    tasksDialog.profileId,
                                    { annualTasks: next },
                                  );
                                }}
                              >
                                Save
                              </button>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={tasksSaving}
                                onClick={() => {
                                  setTasksEditingId(null);
                                  setTasksEditLabel("");
                                  setTasksEditDueDate("");
                                }}
                              >
                                Cancel
                              </button>
                            </>
                          ) : (
                            <>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={tasksSaving}
                                onClick={() => {
                                  setTasksEditingId(task.id);
                                  setTasksEditLabel(task.label);
                                  setTasksEditDueDate(task.dueDate);
                                }}
                              >
                                Edit
                              </button>
                              <button
                                type="button"
                                className="table-link-button one-time-tasks-delete"
                                disabled={tasksSaving}
                                onClick={() => {
                                  if (!tasksDialog) {
                                    return;
                                  }
                                  void persistProfileTasks(
                                    tasksDialog.profileId,
                                    {
                                      annualTasks: removeAnnualTask(
                                        annualDraft,
                                        task.id,
                                      ),
                                    },
                                  );
                                }}
                              >
                                Delete
                              </button>
                            </>
                          )}
                        </div>
                      ) : (
                        <span className="annual-tasks-col-actions" />
                      )}
                    </div>
                    {task.completed ? (
                      <p className="one-time-tasks-meta">
                        Completed{" "}
                        {task.completedAt
                          ? new Date(task.completedAt).toLocaleString()
                          : ""}
                        {task.completedByName
                          ? ` · ${task.completedByName}`
                          : ""}
                      </p>
                    ) : null}
                  </li>
                );
              })}
            </EngagementTasksDialog>,
            document.body,
          )
        : null}

      {tasksDialog?.kind === "monthly"
        ? createPortal(
            <EngagementTasksDialog
              title="Monthly Tasks"
              titleId="ra-monthly-tasks-title"
              statusLabel={monthlyTasksStatusLabel(monthlyDraft) || "Pending"}
              itemCount={monthlyDraft.length}
              saving={tasksSaving}
              error={tasksError}
              canEdit={canEdit}
              onClose={closeTasksDialog}
              intro={
                <>
                  Monthly-billed services from the linked proposal. Empty when
                  no proposal. Defaults cannot be edited or deleted.
                </>
              }
              emptyHint={
                <>
                  <p>No monthly tasks yet.</p>
                  <p>
                    Items appear when the linked proposal includes services
                    billed monthly.
                  </p>
                </>
              }
              addRow={
                <div className="one-time-tasks-add">
                  <label className="field one-time-tasks-add-field">
                    <span className="field-label">Add another task</span>
                    <input
                      className="field-input"
                      value={tasksNewLabel}
                      disabled={tasksSaving}
                      placeholder="Type a new monthly task…"
                      onChange={(event) => setTasksNewLabel(event.target.value)}
                    />
                  </label>
                  <button
                    type="button"
                    className="button secondary-solid"
                    disabled={tasksSaving || !tasksNewLabel.trim()}
                    onClick={() => {
                      const label = tasksNewLabel.trim();
                      if (!label || !tasksDialog) {
                        return;
                      }
                      setTasksNewLabel("");
                      void persistProfileTasks(tasksDialog.profileId, {
                        monthlyTasks: [
                          ...monthlyDraft,
                          createManualMonthlyTask(label),
                        ],
                      });
                    }}
                  >
                    Add
                  </button>
                </div>
              }
            >
              {monthlyDraft.map((task) => {
                const manual = isManualMonthlyTask(task);
                const editing = tasksEditingId === task.id;
                return (
                  <li
                    key={task.id}
                    className={
                      task.completed
                        ? "one-time-tasks-item one-time-tasks-item--done"
                        : "one-time-tasks-item"
                    }
                  >
                    <div className="one-time-tasks-item-row">
                      <label className="one-time-tasks-check">
                        <input
                          type="checkbox"
                          checked={task.completed}
                          disabled={!canEdit || tasksSaving}
                          onChange={(event) => {
                            if (!tasksDialog) {
                              return;
                            }
                            void persistProfileTasks(tasksDialog.profileId, {
                              monthlyTasks: toggleMonthlyTaskCompletion(
                                monthlyDraft,
                                task.id,
                                event.target.checked,
                                actor,
                              ),
                            });
                          }}
                        />
                        {editing ? (
                          <input
                            className="field-input one-time-tasks-edit-input"
                            value={tasksEditLabel}
                            disabled={tasksSaving}
                            autoFocus
                            onChange={(event) =>
                              setTasksEditLabel(event.target.value)
                            }
                          />
                        ) : (
                          <span className="one-time-tasks-label">
                            {task.label}
                            {isProtectedMonthlyTask(task) ? (
                              <span className="one-time-tasks-lock">
                                {" "}
                                · Default
                              </span>
                            ) : null}
                          </span>
                        )}
                      </label>
                      {canEdit && manual ? (
                        <div className="one-time-tasks-item-actions">
                          {editing ? (
                            <>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={
                                  tasksSaving || !tasksEditLabel.trim()
                                }
                                onClick={() => {
                                  if (!tasksDialog) {
                                    return;
                                  }
                                  const next = updateMonthlyTaskLabel(
                                    monthlyDraft,
                                    task.id,
                                    tasksEditLabel,
                                  );
                                  setTasksEditingId(null);
                                  setTasksEditLabel("");
                                  void persistProfileTasks(
                                    tasksDialog.profileId,
                                    { monthlyTasks: next },
                                  );
                                }}
                              >
                                Save
                              </button>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={tasksSaving}
                                onClick={() => {
                                  setTasksEditingId(null);
                                  setTasksEditLabel("");
                                }}
                              >
                                Cancel
                              </button>
                            </>
                          ) : (
                            <>
                              <button
                                type="button"
                                className="table-link-button"
                                disabled={tasksSaving}
                                onClick={() => {
                                  setTasksEditingId(task.id);
                                  setTasksEditLabel(task.label);
                                }}
                              >
                                Edit
                              </button>
                              <button
                                type="button"
                                className="table-link-button one-time-tasks-delete"
                                disabled={tasksSaving}
                                onClick={() => {
                                  if (!tasksDialog) {
                                    return;
                                  }
                                  void persistProfileTasks(
                                    tasksDialog.profileId,
                                    {
                                      monthlyTasks: removeMonthlyTask(
                                        monthlyDraft,
                                        task.id,
                                      ),
                                    },
                                  );
                                }}
                              >
                                Delete
                              </button>
                            </>
                          )}
                        </div>
                      ) : null}
                    </div>
                    {task.completed ? (
                      <p className="one-time-tasks-meta">
                        Completed{" "}
                        {task.completedAt
                          ? new Date(task.completedAt).toLocaleString()
                          : ""}
                        {task.completedByName
                          ? ` · ${task.completedByName}`
                          : ""}
                      </p>
                    ) : null}
                  </li>
                );
              })}
            </EngagementTasksDialog>,
            document.body,
          )
        : null}
    </section>
  );
}
