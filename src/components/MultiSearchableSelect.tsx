"use client";

import { useEffect, useId, useMemo, useRef, useState } from "react";
import { formatMultiValue } from "@/lib/resource-clients";

type MultiSearchableSelectProps = {
  label: string;
  values: string[];
  options: readonly string[];
  placeholder?: string;
  required?: boolean;
  allowCreate?: boolean;
  onChange: (values: string[]) => void;
  onCreateOption?: (label: string) => void | Promise<void>;
};

export function MultiSearchableSelect({
  label,
  values,
  options,
  placeholder,
  required = false,
  allowCreate = false,
  onChange,
  onCreateOption,
}: MultiSearchableSelectProps) {
  const listId = useId();
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [creating, setCreating] = useState(false);

  useEffect(() => {
    function handlePointerDown(event: MouseEvent) {
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
    if (!q) {
      return options;
    }
    return options.filter((option) => option.toLowerCase().includes(q));
  }, [options, query]);

  const trimmedQuery = query.trim();
  const exactExists = options.some(
    (option) => option.trim().toLowerCase() === trimmedQuery.toLowerCase(),
  );
  const canCreate =
    allowCreate && Boolean(trimmedQuery) && !exactExists && !creating;

  function toggleOption(option: string) {
    if (values.includes(option)) {
      onChange(values.filter((value) => value !== option));
      return;
    }
    onChange([...values, option]);
  }

  function removeOption(option: string) {
    onChange(values.filter((value) => value !== option));
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
    <div className="field searchable-select" ref={rootRef}>
      <span className="field-label">
        {label}
        {required ? <span className="field-required"> *</span> : null}
      </span>

      {values.length > 0 ? (
        <div className="multi-select-chips">
          {values.map((value) => (
            <button
              key={value}
              type="button"
              className="multi-select-chip"
              onClick={() => removeOption(value)}
              title={`Remove ${value}`}
            >
              {value} <span aria-hidden="true">×</span>
            </button>
          ))}
        </div>
      ) : null}

      {values.length > 0 ? (
        <p className="multi-select-summary">{formatMultiValue(values)}</p>
      ) : null}

      <div className="searchable-select-control">
        <input
          className="field-input"
          role="combobox"
          aria-expanded={open}
          aria-controls={listId}
          aria-autocomplete="list"
          placeholder={placeholder ?? "Search/Select from the list"}
          value={query}
          onFocus={() => setOpen(true)}
          onChange={(event) => {
            setQuery(event.target.value);
            setOpen(true);
          }}
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

        {open ? (
          <ul id={listId} className="searchable-select-list" role="listbox">
            {filtered.map((option) => {
              const selected = values.includes(option);
              return (
                <li key={option}>
                  <button
                    type="button"
                    role="option"
                    aria-selected={selected}
                    className={
                      selected
                        ? "searchable-select-option active"
                        : "searchable-select-option"
                    }
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => toggleOption(option)}
                  >
                    <span className="multi-select-check">
                      {selected ? "✓" : ""}
                    </span>
                    {option}
                  </button>
                </li>
              );
            })}

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
