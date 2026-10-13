"use client";

import { useEffect, useId, useRef, useState, type ReactNode } from "react";

export type ToolbarFilterDef = {
  id: string;
  label: string;
  options: string[];
};

export type ToolbarSortOption = {
  id: string;
  label: string;
};

type DataListToolbarProps = {
  search: string;
  onSearchChange: (value: string) => void;
  searchPlaceholder: string;
  sort: string;
  onSortChange: (value: string) => void;
  sortOptions: ToolbarSortOption[];
  filters: ToolbarFilterDef[];
  activeFilters: Record<string, string[]>;
  onFilterChange: (filterId: string, values: string[]) => void;
  actions?: ReactNode;
};

export function DataListToolbar({
  search,
  onSearchChange,
  searchPlaceholder,
  sort,
  onSortChange,
  sortOptions,
  filters,
  activeFilters,
  onFilterChange,
  actions,
}: DataListToolbarProps) {
  const [openFilterId, setOpenFilterId] = useState<string | null>(null);
  const [filterSearch, setFilterSearch] = useState("");
  const openFilterRef = useRef<HTMLDivElement | null>(null);
  const listId = useId();

  useEffect(() => {
    if (!openFilterId) {
      return;
    }

    function handlePointerDown(event: MouseEvent) {
      const target = event.target as Node;
      if (openFilterRef.current?.contains(target)) {
        return;
      }
      setOpenFilterId(null);
      setFilterSearch("");
    }

    function handleKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        setOpenFilterId(null);
        setFilterSearch("");
      }
    }

    document.addEventListener("mousedown", handlePointerDown);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("mousedown", handlePointerDown);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [openFilterId]);

  function openFilter(filterId: string) {
    setOpenFilterId((current) => {
      if (current === filterId) {
        setFilterSearch("");
        return null;
      }
      setFilterSearch("");
      return filterId;
    });
  }

  function toggleFilterValue(filterId: string, value: string) {
    const current = activeFilters[filterId] ?? [];
    const next = current.includes(value)
      ? current.filter((item) => item !== value)
      : [...current, value];
    onFilterChange(filterId, next);
  }

  function clearFilter(filterId: string) {
    onFilterChange(filterId, []);
  }

  return (
    <div className="list-toolbar">
      <div className="list-toolbar-row list-toolbar-row--search">
        <label className="list-toolbar-field list-toolbar-field--search">
          <span className="list-toolbar-label">Search</span>
          <input
            className="field-input"
            type="search"
            value={search}
            placeholder={searchPlaceholder}
            onChange={(event) => onSearchChange(event.target.value)}
          />
        </label>
        <label className="list-toolbar-field list-toolbar-field--sort">
          <span className="list-toolbar-label">Sort</span>
          <select
            className="field-input field-select"
            value={sort}
            onChange={(event) => onSortChange(event.target.value)}
          >
            {sortOptions.map((option) => (
              <option key={option.id} value={option.id}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
      </div>

      <div className="list-toolbar-row list-toolbar-row--filters">
        <div className="list-toolbar-filters">
          <span className="list-toolbar-label">Filters</span>
          <div className="list-filter-pills">
            {filters.map((filter) => {
              const selected = activeFilters[filter.id] ?? [];
              const isOpen = openFilterId === filter.id;
              const active = selected.length > 0;
              const query = filterSearch.trim().toLowerCase();
              const visibleOptions = query
                ? filter.options.filter((option) =>
                    option.toLowerCase().includes(query),
                  )
                : filter.options;
              return (
                <div
                  key={filter.id}
                  className="list-filter-pill-wrap"
                  ref={isOpen ? openFilterRef : undefined}
                >
                  <button
                    type="button"
                    className={
                      active || isOpen
                        ? "list-filter-pill active"
                        : "list-filter-pill"
                    }
                    aria-expanded={isOpen}
                    aria-controls={`${listId}-${filter.id}`}
                    onClick={() => openFilter(filter.id)}
                  >
                    {filter.label}
                    {active ? (
                      <span className="list-filter-count">{selected.length}</span>
                    ) : null}
                  </button>
                  {isOpen ? (
                    <div
                      id={`${listId}-${filter.id}`}
                      className="list-filter-menu"
                      role="listbox"
                      aria-label={`${filter.label} filter`}
                    >
                      {filter.options.length > 0 ? (
                        <input
                          className="field-input list-filter-search"
                          type="search"
                          value={filterSearch}
                          placeholder={`Search ${filter.label.toLowerCase()}...`}
                          autoFocus
                          onChange={(event) =>
                            setFilterSearch(event.target.value)
                          }
                          onClick={(event) => event.stopPropagation()}
                          onKeyDown={(event) => event.stopPropagation()}
                        />
                      ) : null}
                      {filter.options.length === 0 ? (
                        <p className="list-filter-empty">No values yet</p>
                      ) : visibleOptions.length === 0 ? (
                        <p className="list-filter-empty">No matches</p>
                      ) : (
                        <div className="list-filter-options">
                          {visibleOptions.map((option) => {
                            const checked = selected.includes(option);
                            return (
                              <label
                                key={option}
                                className="list-filter-option"
                              >
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() =>
                                    toggleFilterValue(filter.id, option)
                                  }
                                />
                                <span>{option}</span>
                              </label>
                            );
                          })}
                        </div>
                      )}
                      {active ? (
                        <button
                          type="button"
                          className="list-filter-clear"
                          onClick={() => clearFilter(filter.id)}
                        >
                          Clear
                        </button>
                      ) : null}
                    </div>
                  ) : null}
                </div>
              );
            })}
          </div>
        </div>

        {actions ? (
          <div className="list-toolbar-aside">
            <div className="panel-title-actions">{actions}</div>
          </div>
        ) : null}
      </div>
    </div>
  );
}

export function uniqueSortedValues(values: string[]): string[] {
  return [...new Set(values.map((value) => value.trim()).filter(Boolean))].sort(
    (a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }),
  );
}

export function downloadCsv(filename: string, rows: string[][]) {
  const escape = (value: string) => `"${value.replace(/"/g, '""')}"`;
  const csv = rows.map((row) => row.map(escape).join(",")).join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

/** Parse a CSV string into rows of cells (supports quoted commas). */
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  const source = text.replace(/^\uFEFF/, "");
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    const next = source[index + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        cell += '"';
        index += 1;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        cell += char;
      }
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      continue;
    }
    if (char === ",") {
      row.push(cell);
      cell = "";
      continue;
    }
    if (char === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }
    if (char === "\r") {
      continue;
    }
    cell += char;
  }

  if (cell.length > 0 || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }

  return rows.filter((entry) => entry.some((value) => value.trim()));
}
