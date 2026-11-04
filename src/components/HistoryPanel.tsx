"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { useAccess } from "@/auth/accessContext";
import {
  DataListToolbar,
  downloadCsv,
  parseSpreadsheetFile,
  SPREADSHEET_IMPORT_ACCEPT,
  uniqueSortedValues,
} from "@/components/DataListToolbar";
import {
  actionLabel,
  entityTypeLabel,
  formatHistoryDate,
  formatHistoryTime,
  type HistoryAction,
  type HistoryEntityType,
  type HistoryEntry,
} from "@/lib/history";
import { portalFetch } from "@/lib/portal-api";

const HISTORY_SORT_OPTIONS = [
  { id: "newest", label: "Newest first" },
  { id: "oldest", label: "Oldest first" },
];

export function HistoryPanel() {
  const { canEdit } = useAccess();
  const [entries, setEntries] = useState<HistoryEntry[]>([]);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState("newest");
  const [clearing, setClearing] = useState(false);
  const [importing, setImporting] = useState(false);
  const importInputRef = useRef<HTMLInputElement>(null);
  const [activeFilters, setActiveFilters] = useState<Record<string, string[]>>(
    {},
  );

  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const response = await fetch("/api/history", { cache: "no-store" });
        const data = (await response.json()) as {
          entries?: HistoryEntry[];
          error?: string;
        };
        if (!response.ok) {
          throw new Error(data.error || "Failed to load history.");
        }
        if (!cancelled) {
          setEntries(data.entries ?? []);
          setError(null);
        }
      } catch (loadError) {
        if (!cancelled) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load history.",
          );
        }
      } finally {
        if (!cancelled) {
          setReady(true);
        }
      }
    }

    void load();
    return () => {
      cancelled = true;
    };
  }, []);

  const toolbarFilters = useMemo(
    () => [
      {
        id: "employee",
        label: "Employee",
        options: uniqueSortedValues(
          entries.map((entry) => entry.employeeName),
        ),
      },
      {
        id: "project",
        label: "Client / Project",
        options: uniqueSortedValues(entries.map((entry) => entry.project)),
      },
      {
        id: "area",
        label: "Area",
        options: uniqueSortedValues(
          entries.map((entry) => entityTypeLabel(entry.entityType)),
        ),
      },
      {
        id: "action",
        label: "Action",
        options: uniqueSortedValues(
          entries.map((entry) => actionLabel(entry.action)),
        ),
      },
    ],
    [entries],
  );

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    let rows = entries.filter((entry) => {
      if (query) {
        const haystack = [
          entry.employeeName,
          entry.employeeEmail,
          entry.project,
          entry.summary,
          entry.details,
          entry.comment,
          entry.entityLabel,
          actionLabel(entry.action),
          entityTypeLabel(entry.entityType),
        ]
          .join(" ")
          .toLowerCase();
        if (!haystack.includes(query)) {
          return false;
        }
      }

      const employeeFilter = activeFilters.employee ?? [];
      if (
        employeeFilter.length > 0 &&
        !employeeFilter.includes(entry.employeeName.trim())
      ) {
        return false;
      }

      const projectFilter = activeFilters.project ?? [];
      if (
        projectFilter.length > 0 &&
        !projectFilter.includes(entry.project.trim())
      ) {
        return false;
      }

      const areaFilter = activeFilters.area ?? [];
      if (
        areaFilter.length > 0 &&
        !areaFilter.includes(entityTypeLabel(entry.entityType))
      ) {
        return false;
      }

      const actionFilter = activeFilters.action ?? [];
      if (
        actionFilter.length > 0 &&
        !actionFilter.includes(actionLabel(entry.action))
      ) {
        return false;
      }

      return true;
    });

    rows = [...rows].sort((a, b) => {
      if (sort === "oldest") {
        return (
          new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime()
        );
      }
      return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
    });

    return rows;
  }, [entries, search, activeFilters, sort]);

  function handleFilterChange(filterId: string, values: string[]) {
    setActiveFilters((current) => ({ ...current, [filterId]: values }));
  }

  function handleExport() {
    const header = [
      "Date",
      "Time",
      "Employee",
      "Employee Email",
      "Client / Project",
      "Area",
      "Action",
      "Summary",
      "Details",
      "Comments",
    ];
    const rows = filtered.map((entry) => [
      formatHistoryDate(entry.createdAt),
      formatHistoryTime(entry.createdAt),
      entry.employeeName,
      entry.employeeEmail,
      entry.project,
      entityTypeLabel(entry.entityType),
      actionLabel(entry.action),
      entry.summary,
      entry.details,
      entry.comment,
    ]);
    downloadCsv("history.csv", [header, ...rows]);
  }

  function openImportPicker() {
    importInputRef.current?.click();
  }

  function parseEntityType(raw: string): HistoryEntityType {
    const value = raw.trim().toLowerCase();
    if (
      value === "clientprofile" ||
      value === "client profile" ||
      value.includes("client profile")
    ) {
      return "clientProfile";
    }
    if (
      value === "lookupdata" ||
      value === "lookup data" ||
      value.includes("lookup")
    ) {
      return "lookupData";
    }
    return "resourceAllocation";
  }

  function parseAction(raw: string): HistoryAction {
    const value = raw.trim().toLowerCase();
    if (value.startsWith("creat")) {
      return "created";
    }
    if (value.startsWith("delet")) {
      return "deleted";
    }
    return "updated";
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
      const indexOf = (labels: string[]) =>
        headers.findIndex((header) => labels.includes(header));

      const dateIdx = indexOf(["date"]);
      const timeIdx = indexOf(["time"]);
      const employeeIdx = indexOf(["employee", "employee name"]);
      const emailIdx = indexOf(["employee email", "email"]);
      const projectIdx = indexOf(["client / project", "client", "project"]);
      const areaIdx = indexOf(["area", "entity type"]);
      const actionIdx = indexOf(["action"]);
      const summaryIdx = indexOf(["summary"]);
      const detailsIdx = indexOf(["details"]);
      const commentIdx = indexOf(["comments", "comment"]);

      if (employeeIdx < 0 && summaryIdx < 0) {
        throw new Error(
          "No recognized columns. Export first and use those headers.",
        );
      }

      const imported: HistoryEntry[] = [];
      for (const row of rows.slice(1)) {
        const employeeName = (row[employeeIdx] ?? "").trim();
        const summary = (row[summaryIdx] ?? "").trim();
        if (!employeeName && !summary) {
          continue;
        }

        const datePart = (row[dateIdx] ?? "").trim();
        const timePart = (row[timeIdx] ?? "").trim();
        const combined = [datePart, timePart].filter(Boolean).join(" ");
        const parsedDate = combined ? new Date(combined) : null;
        const createdAt =
          parsedDate && !Number.isNaN(parsedDate.getTime())
            ? parsedDate.toISOString()
            : new Date().toISOString();

        imported.push({
          id: `hist-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          createdAt,
          employeeName: employeeName || "Unknown",
          employeeEmail: (row[emailIdx] ?? "").trim(),
          project: (row[projectIdx] ?? "").trim(),
          entityType: parseEntityType(row[areaIdx] ?? ""),
          entityLabel: (row[projectIdx] ?? "").trim(),
          action: parseAction(row[actionIdx] ?? ""),
          summary: summary || "Imported history entry",
          details: (row[detailsIdx] ?? "").trim(),
          comment: (row[commentIdx] ?? "").trim(),
        });
      }

      if (imported.length === 0) {
        throw new Error("No valid history rows found.");
      }

      let added = 0;
      for (const entry of imported) {
        const response = await portalFetch("/api/history", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            employeeName: entry.employeeName,
            employeeEmail: entry.employeeEmail,
            project: entry.project,
            entityType: entry.entityType,
            entityLabel: entry.entityLabel,
            action: entry.action,
            summary: entry.summary,
            details: entry.details,
            comment: entry.comment,
            createdAt: entry.createdAt,
          }),
        });
        if (!response.ok) {
          const data = (await response.json()) as { error?: string };
          throw new Error(data.error || "Failed to import history row.");
        }
        added += 1;
      }

      const refresh = await fetch("/api/history", { cache: "no-store" });
      const refreshData = (await refresh.json()) as {
        entries?: HistoryEntry[];
      };
      if (refresh.ok) {
        setEntries(refreshData.entries ?? []);
      }
      setError(null);
      window.alert(`Imported ${added} history ${added === 1 ? "entry" : "entries"}.`);
    } catch (importError) {
      setError(
        importError instanceof Error
          ? importError.message
          : "Failed to import history.",
      );
    } finally {
      setImporting(false);
      if (importInputRef.current) {
        importInputRef.current.value = "";
      }
    }
  }

  async function handleClearHistory() {
    if (entries.length === 0) {
      return;
    }
    if (
      !window.confirm(
        `Clear all ${entries.length} history ${entries.length === 1 ? "entry" : "entries"}? This cannot be undone.`,
      )
    ) {
      return;
    }

    setClearing(true);
    setError(null);
    try {
      const response = await portalFetch("/api/history", { method: "DELETE" });
      const data = (await response.json()) as { error?: string };
      if (!response.ok) {
        throw new Error(data.error || "Failed to clear history.");
      }
      setEntries([]);
      setSearch("");
      setActiveFilters({});
    } catch (clearError) {
      setError(
        clearError instanceof Error
          ? clearError.message
          : "Failed to clear history.",
      );
    } finally {
      setClearing(false);
    }
  }

  return (
    <section className="content-panel content-panel--actions">
      <DataListToolbar
        search={search}
        onSearchChange={setSearch}
        searchPlaceholder="Employee, client, action, summary..."
        sort={sort}
        onSortChange={setSort}
        sortOptions={HISTORY_SORT_OPTIONS}
        filters={toolbarFilters}
        activeFilters={activeFilters}
        onFilterChange={handleFilterChange}
        actions={
          <>
            {canEdit ? (
              <button
                type="button"
                className="button danger"
                onClick={() => void handleClearHistory()}
                disabled={entries.length === 0 || clearing}
              >
                {clearing ? "Clearing…" : "Clear"}
              </button>
            ) : null}
            <button
              type="button"
              className="button secondary-solid"
              onClick={handleExport}
              disabled={filtered.length === 0}
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
              </>
            ) : null}
          </>
        }
      />

      {error ? <p className="form-message error">{error}</p> : null}

      <div className="table-wrap">
        {!ready ? (
          <p className="table-empty">Loading…</p>
        ) : entries.length === 0 ? (
          <p className="table-empty">
            No history yet. Create or edit a Client Profile or Resource
            Allocation to start logging.
          </p>
        ) : filtered.length === 0 ? (
          <p className="table-empty">No history matches the current filters.</p>
        ) : (
          <table className="projects-table projects-table--wide">
            <thead>
              <tr>
                <th>Date</th>
                <th>Time</th>
                <th>Employee</th>
                <th>Client / Project</th>
                <th>Area</th>
                <th>Action</th>
                <th>Summary</th>
                <th>Details</th>
                <th>Comments</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((entry) => (
                <tr key={entry.id}>
                  <td>{formatHistoryDate(entry.createdAt)}</td>
                  <td>{formatHistoryTime(entry.createdAt)}</td>
                  <td>
                    <div>{entry.employeeName}</div>
                    {entry.employeeEmail ? (
                      <div className="history-meta">{entry.employeeEmail}</div>
                    ) : null}
                  </td>
                  <td>{entry.project || "—"}</td>
                  <td>{entityTypeLabel(entry.entityType)}</td>
                  <td>{actionLabel(entry.action)}</td>
                  <td>{entry.summary || "—"}</td>
                  <td className="cell-multiline">{entry.details || "—"}</td>
                  <td className="cell-multiline">{entry.comment || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </section>
  );
}
