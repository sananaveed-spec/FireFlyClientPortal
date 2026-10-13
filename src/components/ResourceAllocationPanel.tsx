"use client";

import { useMsal } from "@azure/msal-react";
import { useEffect, useMemo, useState, type FormEvent } from "react";
import { createPortal } from "react-dom";
import { useAccess } from "@/auth/accessContext";
import {
  DataListToolbar,
  downloadCsv,
  uniqueSortedValues,
} from "@/components/DataListToolbar";
import { RowActionsMenu } from "@/components/RowActionsMenu";
import { SearchableChecklist } from "@/components/SearchableChecklist";
import {
  SearchableSelect,
  type SearchableOption,
} from "@/components/SearchableSelect";
import type { ClientProfileEntry } from "@/lib/client-profiles";
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
import {
  bookkeepingCategoryForCountry,
  createEmptyResourceClientForm,
  formatMultiValue,
  resolveMeetingFrequency,
  type ResourceClientEntry,
  type ResourceClientInput,
} from "@/lib/resource-clients";
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

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function loadAll() {
      try {
        const [clientsRes, profilesRes, lookupsRes] = await Promise.all([
          fetch("/api/resource-clients"),
          fetch("/api/client-profiles"),
          fetch("/api/lookup-data", { cache: "no-store" }),
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
  const supervisorOptions = useMemo(
    () => getLookupOptions(lookups, "primaryAssignee"),
    [lookups],
  );
  const bookkeeperOptions = useMemo(
    () => getLookupOptions(lookups, "secondaryAssignee"),
    [lookups],
  );
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
      supervisor: client.supervisor,
      bookkeeperAccountant: [...client.bookkeeperAccountant],
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
      supervisor: form.supervisor.trim(),
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
              <button
                type="button"
                className="button primary"
                onClick={openNewClientForm}
              >
                New Client
              </button>
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
                  <td>{client.supervisor || "—"}</td>
                  <td>
                    {formatMultiValue(client.bookkeeperAccountant) || "—"}
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
                      label="Supervisor"
                      value={form.supervisor}
                      options={supervisorOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(value) => updateForm("supervisor", value)}
                      onCreateOption={(label) =>
                        createSingleLookupValue(
                          "primaryAssignee",
                          "supervisor",
                          label,
                        )
                      }
                    />

                    <SearchableChecklist
                      label="Bookkeeper/Accountant"
                      values={form.bookkeeperAccountant}
                      options={bookkeeperOptions}
                      allowCreate
                      placeholder="Search/Select from the list"
                      onChange={(values) =>
                        updateForm("bookkeeperAccountant", values)
                      }
                      onCreateOption={(label) =>
                        createMultiLookupValue(
                          "secondaryAssignee",
                          "bookkeeperAccountant",
                          label,
                          form.bookkeeperAccountant,
                        )
                      }
                    />

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
    </section>
  );
}
