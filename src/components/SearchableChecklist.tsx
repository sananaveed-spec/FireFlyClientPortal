"use client";

import {
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
  type MouseEvent,
} from "react";
import { formatMultiValue } from "@/lib/resource-clients";

export type ChecklistOption = {
  value: string;
  label: string;
  meta?: string;
  keywords?: string;
};

type SearchableChecklistProps = {
  label: string;
  values: string[];
  options: readonly string[] | readonly ChecklistOption[];
  placeholder?: string;
  allowCreate?: boolean;
  /** When this option is selected, no other options can be selected. */
  exclusiveOption?: string;
  onChange: (values: string[]) => void;
  onCreateOption?: (label: string) => void | Promise<void>;
};

function normalizeOptions(
  options: readonly string[] | readonly ChecklistOption[],
): ChecklistOption[] {
  return options.map((option) =>
    typeof option === "string"
      ? { value: option, label: option }
      : option,
  );
}

function isSameOption(a: string, b: string) {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

export function SearchableChecklist({
  label,
  values,
  options,
  placeholder,
  allowCreate = false,
  exclusiveOption,
  onChange,
  onCreateOption,
}: SearchableChecklistProps) {
  const listId = useId();
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [creating, setCreating] = useState(false);
  const normalized = useMemo(() => normalizeOptions(options), [options]);

  const displayText =
    values.length > 0
      ? formatMultiValue(
          values.map((value) => {
            const match = normalized.find((option) =>
              isSameOption(option.value, value),
            );
            return match?.label ?? value;
          }),
        )
      : (placeholder ?? "Search/Select from the list");

  const exclusiveSelected = Boolean(
    exclusiveOption &&
      values.some((value) => isSameOption(value, exclusiveOption)),
  );

  useEffect(() => {
    function handlePointerDown(event: globalThis.MouseEvent) {
      if (!rootRef.current?.contains(event.target as Node)) {
        setOpen(false);
        setQuery("");
      }
    }

    document.addEventListener("mousedown", handlePointerDown);
    return () => document.removeEventListener("mousedown", handlePointerDown);
  }, []);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    let next = normalized;
    if (exclusiveSelected && exclusiveOption) {
      next = normalized.filter((option) =>
        isSameOption(option.value, exclusiveOption),
      );
    }
    if (!q) {
      return next;
    }
    return next.filter((option) => {
      const haystack =
        `${option.label} ${option.value} ${option.keywords ?? ""}`.toLowerCase();
      return haystack.includes(q);
    });
  }, [normalized, query, exclusiveSelected, exclusiveOption]);

  const trimmedQuery = query.trim();
  const exactExists = normalized.some(
    (option) =>
      option.label.trim().toLowerCase() === trimmedQuery.toLowerCase() ||
      option.value.trim().toLowerCase() === trimmedQuery.toLowerCase(),
  );
  const canCreate =
    allowCreate &&
    Boolean(trimmedQuery) &&
    !exactExists &&
    !creating &&
    !exclusiveSelected;

  function toggleOption(optionValue: string) {
    if (values.some((value) => isSameOption(value, optionValue))) {
      onChange(values.filter((value) => !isSameOption(value, optionValue)));
      return;
    }

    if (exclusiveOption && isSameOption(optionValue, exclusiveOption)) {
      onChange([optionValue]);
      return;
    }

    if (exclusiveSelected) {
      return;
    }

    onChange([...values, optionValue]);
  }

  function clearSelection(event: MouseEvent<HTMLButtonElement>) {
    event.preventDefault();
    event.stopPropagation();
    onChange([]);
  }

  async function createOption() {
    if (!canCreate) {
      return;
    }
    setCreating(true);
    try {
      if (onCreateOption) {
        await onCreateOption(trimmedQuery);
      } else {
        onChange([...values, trimmedQuery]);
      }
      setQuery("");
    } finally {
      setCreating(false);
    }
  }

  return (
    <div className="field searchable-select checklist-dropdown" ref={rootRef}>
      <span className="field-label">{label}</span>

      <div className="searchable-select-control">
        <button
          type="button"
          className={`field-input checklist-dropdown-trigger${values.length > 0 ? " checklist-dropdown-trigger--filled" : ""}`}
          aria-expanded={open}
          aria-controls={listId}
          onClick={() => setOpen((current) => !current)}
        >
          <span className="checklist-dropdown-value">{displayText}</span>
          {values.length === 0 ? (
            <span className="checklist-dropdown-caret" aria-hidden="true">
              {open ? "▴" : "▾"}
            </span>
          ) : null}
        </button>
        {values.length > 0 ? (
          <button
            type="button"
            className="searchable-select-clear"
            aria-label={`Clear ${label}`}
            onClick={clearSelection}
          >
            ×
          </button>
        ) : null}
      </div>

      {open ? (
        <div id={listId} className="checklist-dropdown-panel">
          <input
            className="field-input"
            type="search"
            value={query}
            placeholder="Search/Select from the list"
            autoFocus
            onChange={(event) => setQuery(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Escape") {
                setOpen(false);
                setQuery("");
              }
              if (event.key === "Enter" && canCreate) {
                event.preventDefault();
                void createOption();
              }
            }}
          />
          {exclusiveSelected ? (
            <p className="checklist-hint">
              Clear NIL to see other options.
            </p>
          ) : null}
          <div className="checklist-box" role="group" aria-label={label}>
            {filtered.map((option) => {
              const checked = values.some((value) =>
                isSameOption(value, option.value),
              );
              return (
                <label key={option.value} className="checklist-option">
                  <input
                    type="checkbox"
                    checked={checked}
                    onChange={() => toggleOption(option.value)}
                  />
                  <span>
                    {option.label}
                    {option.meta ? (
                      <span className="checklist-option-meta">
                        {" "}
                        ({option.meta})
                      </span>
                    ) : null}
                  </span>
                </label>
              );
            })}

            {canCreate ? (
              <button
                type="button"
                className="searchable-select-option searchable-select-create"
                onClick={() => void createOption()}
                disabled={creating}
              >
                {creating ? "Adding…" : `Add “${trimmedQuery}”`}
              </button>
            ) : null}

            {filtered.length === 0 && !canCreate ? (
              <p className="table-empty">No matches</p>
            ) : null}
          </div>
        </div>
      ) : null}
    </div>
  );
}
