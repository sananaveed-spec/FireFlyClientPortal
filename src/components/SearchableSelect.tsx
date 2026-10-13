"use client";

import { useEffect, useId, useMemo, useRef, useState } from "react";

export type SearchableOption = {
  value: string;
  label: string;
  /** Extra text used only for search matching (e.g. Unique ID). */
  keywords?: string;
  /** Secondary line shown in the dropdown (e.g. Unique ID). */
  meta?: string;
};

type SearchableSelectProps = {
  label: string;
  value: string;
  options: readonly string[] | readonly SearchableOption[];
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  /** Hide the visible label (still used for aria). */
  hideLabel?: boolean;
  className?: string;
  /** Allow creating a new value when it is not in the list. */
  allowCreate?: boolean;
  onChange: (value: string) => void;
  onCreateOption?: (label: string) => void | Promise<void>;
};

function normalizeOptions(
  options: readonly string[] | readonly SearchableOption[],
): SearchableOption[] {
  return options.map((option) =>
    typeof option === "string"
      ? { value: option, label: option }
      : option,
  );
}

export function SearchableSelect({
  label,
  value,
  options,
  placeholder,
  required = false,
  disabled = false,
  hideLabel = false,
  className,
  allowCreate = false,
  onChange,
  onCreateOption,
}: SearchableSelectProps) {
  const listId = useId();
  const rootRef = useRef<HTMLDivElement | null>(null);
  const normalized = useMemo(() => normalizeOptions(options), [options]);
  const selected = normalized.find((option) => option.value === value);
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState(selected?.label ?? value);
  const [creating, setCreating] = useState(false);

  useEffect(() => {
    const match = normalized.find((option) => option.value === value);
    setQuery(match?.label ?? value);
  }, [value, normalized]);

  useEffect(() => {
    function handlePointerDown(event: MouseEvent) {
      if (!rootRef.current?.contains(event.target as Node)) {
        setOpen(false);
        const match = normalized.find((option) => option.value === value);
        setQuery(match?.label ?? value);
      }
    }

    document.addEventListener("mousedown", handlePointerDown);
    return () => document.removeEventListener("mousedown", handlePointerDown);
  }, [value, normalized]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) {
      return normalized;
    }
    return normalized.filter((option) => {
      const haystack = [
        option.label,
        option.value,
        option.keywords ?? "",
        option.meta ?? "",
      ]
        .join(" ")
        .toLowerCase();
      return haystack.includes(q);
    });
  }, [normalized, query]);

  const trimmedQuery = query.trim();
  const exactExists = normalized.some(
    (option) => option.label.trim().toLowerCase() === trimmedQuery.toLowerCase(),
  );
  const canCreate =
    allowCreate && Boolean(trimmedQuery) && !exactExists && !disabled;

  function selectOption(option: SearchableOption) {
    if (disabled) {
      return;
    }
    onChange(option.value);
    setQuery(option.label);
    setOpen(false);
  }

  async function createOption() {
    if (!canCreate || creating) {
      return;
    }
    setCreating(true);
    try {
      if (onCreateOption) {
        await onCreateOption(trimmedQuery);
      } else {
        onChange(trimmedQuery);
      }
      setQuery(trimmedQuery);
      setOpen(false);
    } finally {
      setCreating(false);
    }
  }

  function clearSelection() {
    if (disabled) {
      return;
    }
    onChange("");
    setQuery("");
    setOpen(true);
  }

  return (
    <div
      className={`field searchable-select${
        disabled ? " searchable-select--disabled" : ""
      }${hideLabel ? " searchable-select--no-label" : ""}${
        className ? ` ${className}` : ""
      }`}
      ref={rootRef}
    >
      {hideLabel ? (
        <span className="sr-only">
          {label}
          {required ? " (required)" : ""}
        </span>
      ) : (
        <span className="field-label">
          {label}
          {required ? <span className="field-required"> *</span> : null}
        </span>
      )}
      <div className="searchable-select-control">
        <input
          className="field-input"
          role="combobox"
          aria-expanded={open}
          aria-controls={listId}
          aria-autocomplete="list"
          aria-label={hideLabel ? label : undefined}
          placeholder={placeholder ?? "Search/Select from the list"}
          value={query}
          disabled={disabled}
          readOnly={disabled}
          onFocus={() => {
            if (!disabled) {
              setOpen(true);
            }
          }}
          onChange={(event) => {
            if (disabled) {
              return;
            }
            setQuery(event.target.value);
            setOpen(true);
            if (value) {
              const match = normalized.find((option) => option.value === value);
              if (event.target.value !== (match?.label ?? value)) {
                onChange("");
              }
            }
          }}
          onKeyDown={(event) => {
            if (disabled) {
              return;
            }
            if (event.key === "Escape") {
              setOpen(false);
              const match = normalized.find((option) => option.value === value);
              setQuery(match?.label ?? value);
            }
            if (event.key === "Enter") {
              event.preventDefault();
              if (filtered.length === 1 && !canCreate) {
                selectOption(filtered[0]);
                return;
              }
              if (canCreate) {
                void createOption();
              }
            }
          }}
        />
        {value && !disabled ? (
          <button
            type="button"
            className="searchable-select-clear"
            aria-label={`Clear ${label}`}
            onClick={clearSelection}
          >
            ×
          </button>
        ) : null}

        {open && !disabled ? (
          <ul id={listId} className="searchable-select-list" role="listbox">
            {filtered.map((option) => (
              <li key={`${option.value}::${option.label}`}>
                <button
                  type="button"
                  role="option"
                  aria-selected={option.value === value}
                  className={
                    option.value === value
                      ? "searchable-select-option active"
                      : "searchable-select-option"
                  }
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => selectOption(option)}
                >
                  <span className="searchable-select-option-main">
                    {option.label}
                  </span>
                  {option.meta ? (
                    <span className="searchable-select-option-meta">
                      {option.meta}
                    </span>
                  ) : null}
                </button>
              </li>
            ))}

            {canCreate ? (
              <li>
                <button
                  type="button"
                  className="searchable-select-option searchable-select-create"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => void createOption()}
                  disabled={creating}
                >
                  {creating ? "Adding…" : `Add “${trimmedQuery}”`}
                </button>
              </li>
            ) : null}

            {filtered.length === 0 && !canCreate ? (
              <li className="searchable-select-empty">No matches</li>
            ) : null}
          </ul>
        ) : null}
      </div>
    </div>
  );
}
