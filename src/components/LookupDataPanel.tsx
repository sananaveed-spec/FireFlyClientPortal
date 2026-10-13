"use client";

import { useMsal } from "@azure/msal-react";
import { useEffect, useMemo, useState, type FormEvent } from "react";
import { createPortal } from "react-dom";
import { useAccess } from "@/auth/accessContext";
import {
  getActorFromAccount,
  logPortalAction,
} from "@/lib/history-client";
import { EmailTemplatesPanel } from "@/components/EmailTemplatesPanel";
import { HolidaysPanel } from "@/components/HolidaysPanel";
import {
  createLookupItem,
  type LookupDataStore,
  type LookupFieldId,
  type LookupItem,
} from "@/lib/lookup-data";
import { portalFetch } from "@/lib/portal-api";

type LookupSection = "dropdown" | "holidays" | "emailTemplates";

type EditorState = {
  mode: "add" | "edit";
  fieldId: LookupFieldId;
  itemId?: string;
  label: string;
};

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

export function LookupDataPanel() {
  const { accounts } = useMsal();
  const { canEdit } = useAccess();
  const actor = getActorFromAccount(accounts[0]);
  const [section, setSection] = useState<LookupSection>("emailTemplates");
  const [store, setStore] = useState<LookupDataStore>({ fields: [] });
  const [activeFieldId, setActiveFieldId] =
    useState<LookupFieldId>("industry");
  const [parentSearch, setParentSearch] = useState("");
  const [editor, setEditor] = useState<EditorState | null>(null);
  const [ready, setReady] = useState(false);
  const [mounted, setMounted] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const response = await fetch("/api/lookup-data", { cache: "no-store" });
        const data = (await response.json()) as LookupDataStore & {
          error?: string;
        };
        if (!response.ok) {
          throw new Error(data.error || "Failed to load lookup data.");
        }
        if (!cancelled) {
          setStore({ fields: data.fields ?? [] });
          if (data.fields?.[0]?.id) {
            setActiveFieldId(data.fields[0].id);
          }
          setError(null);
        }
      } catch (loadError) {
        if (!cancelled) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load lookup data.",
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

  const filteredFields = useMemo(() => {
    const query = parentSearch.trim().toLowerCase();
    if (!query) {
      return store.fields;
    }
    return store.fields.filter((field) =>
      field.name.toLowerCase().includes(query),
    );
  }, [store.fields, parentSearch]);

  const activeField = useMemo(
    () =>
      store.fields.find((field) => field.id === activeFieldId) ??
      store.fields[0] ??
      null,
    [store.fields, activeFieldId],
  );

  useEffect(() => {
    if (
      filteredFields.length > 0 &&
      !filteredFields.some((field) => field.id === activeFieldId)
    ) {
      setActiveFieldId(filteredFields[0].id);
    }
  }, [filteredFields, activeFieldId]);

  async function saveStore(next: LookupDataStore) {
    setSaving(true);
    setError(null);
    try {
      const saved = await persistLookupData(next);
      setStore({ fields: saved.fields ?? [] });
      setEditor(null);
    } catch (saveError) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save lookup data.",
      );
    } finally {
      setSaving(false);
    }
  }

  function updateFieldItems(
    fieldId: LookupFieldId,
    updater: (items: LookupItem[]) => LookupItem[],
  ) {
    const next: LookupDataStore = {
      fields: store.fields.map((field) =>
        field.id === fieldId
          ? { ...field, items: updater(field.items) }
          : field,
      ),
    };
    void saveStore(next);
  }

  function openAddItem() {
    if (!activeField) {
      return;
    }
    setError(null);
    setEditor({
      mode: "add",
      fieldId: activeField.id,
      label: "",
    });
  }

  function openEditItem(item: LookupItem) {
    if (!activeField) {
      return;
    }
    setError(null);
    setEditor({
      mode: "edit",
      fieldId: activeField.id,
      itemId: item.id,
      label: item.label,
    });
  }

  function deleteItem(item: LookupItem) {
    if (!activeField) {
      return;
    }
    if (!window.confirm(`Delete "${item.label}"?`)) {
      return;
    }
    updateFieldItems(activeField.id, (items) =>
      items.filter((entry) => entry.id !== item.id),
    );
    void logPortalAction({
      actor,
      project: activeField.name,
      entityType: "lookupData",
      entityLabel: `${activeField.name} / ${item.label}`,
      action: "deleted",
      summary: `Lookup item deleted: ${item.label}`,
      details: `Field: ${activeField.name}\nItem: ${item.label}`,
      comment: "Lookup sub-item removed.",
    });
  }

  function closeEditor() {
    if (saving) {
      return;
    }
    setEditor(null);
  }

  function submitEditor(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!editor) {
      return;
    }

    const label = editor.label.trim();
    if (!label) {
      setError("Label is required.");
      return;
    }

    const fieldName =
      store.fields.find((field) => field.id === editor.fieldId)?.name ??
      editor.fieldId;

    if (editor.mode === "add") {
      const item = createLookupItem(label);
      updateFieldItems(editor.fieldId, (items) => [...items, item]);
      void logPortalAction({
        actor,
        project: fieldName,
        entityType: "lookupData",
        entityLabel: `${fieldName} / ${label}`,
        action: "created",
        summary: `Lookup item created: ${label}`,
        details: `Field: ${fieldName}\nItem: ${label}`,
        comment: "New lookup sub-item added.",
      });
      return;
    }

    const previous =
      store.fields
        .find((field) => field.id === editor.fieldId)
        ?.items.find((item) => item.id === editor.itemId)?.label ?? "";

    updateFieldItems(editor.fieldId, (items) =>
      items.map((item) =>
        item.id === editor.itemId ? { ...item, label } : item,
      ),
    );
    void logPortalAction({
      actor,
      project: fieldName,
      entityType: "lookupData",
      entityLabel: `${fieldName} / ${label}`,
      action: "updated",
      summary: `Lookup item updated: ${label}`,
      details: `Field: ${fieldName}\nLabel: "${previous || "—"}" → "${label}"`,
      comment: "Lookup sub-item renamed.",
    });
  }

  return (
    <section className="content-panel content-panel--actions">
      <div className="lookup-section-tabs" role="tablist" aria-label="Lookup sections">
        <button
          type="button"
          role="tab"
          aria-selected={section === "emailTemplates"}
          className={
            section === "emailTemplates"
              ? "lookup-section-tab active"
              : "lookup-section-tab"
          }
          onClick={() => setSection("emailTemplates")}
        >
          Email Templates
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={section === "holidays"}
          className={
            section === "holidays"
              ? "lookup-section-tab active"
              : "lookup-section-tab"
          }
          onClick={() => setSection("holidays")}
        >
          Holidays
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={section === "dropdown"}
          className={
            section === "dropdown"
              ? "lookup-section-tab active"
              : "lookup-section-tab"
          }
          onClick={() => setSection("dropdown")}
        >
          Dropdown
        </button>
      </div>

      {section === "holidays" ? (
        <HolidaysPanel />
      ) : section === "emailTemplates" ? (
        <EmailTemplatesPanel />
      ) : (
        <>
          <div className="panel-title-row lookup-toolbar">
            <label className="lookup-parent-search">
              <span className="list-toolbar-label">Search parent</span>
              <input
                className="field-input"
                type="search"
                value={parentSearch}
                placeholder="Search parent item name..."
                onChange={(event) => setParentSearch(event.target.value)}
              />
            </label>
            {canEdit ? (
              <button
                type="button"
                className="button primary"
                onClick={openAddItem}
                disabled={!activeField || saving}
              >
                Add sub-item
              </button>
            ) : null}
          </div>

          {error && !editor ? (
            <p className="form-message error">{error}</p>
          ) : null}

          {!ready ? (
            <p className="table-empty">Loading…</p>
          ) : (
            <div className="lookup-layout">
              <aside className="lookup-fields" aria-label="Lookup fields">
                {filteredFields.length === 0 ? (
                  <p className="lookup-hint">No parent matches.</p>
                ) : (
                  filteredFields.map((field) => (
                    <button
                      key={field.id}
                      type="button"
                      className={
                        activeField?.id === field.id
                          ? "lookup-field-link active"
                          : "lookup-field-link"
                      }
                      onClick={() => setActiveFieldId(field.id)}
                    >
                      <span>{field.name}</span>
                      <span className="lookup-field-count">
                        {field.items.length}
                      </span>
                    </button>
                  ))
                )}
              </aside>

              <div className="lookup-items">
                {activeField ? (
                  <>
                    <div className="panel-title-row">
                      <div>
                        <p className="lookup-parent-label">Parent</p>
                        <h3 className="section-title">{activeField.name}</h3>
                      </div>
                      <p className="lookup-hint">
                        Add, edit, or delete sub-items for this dropdown.
                        Sub-items cannot have children.
                      </p>
                    </div>

                    {activeField.items.length === 0 ? (
                      <p className="table-empty">
                        No sub-items yet. Click Add sub-item to create one.
                      </p>
                    ) : (
                      <div className="lookup-item-list">
                        {activeField.items.map((item) => (
                          <div
                            key={item.id}
                            className="lookup-child-row lookup-sub-row"
                          >
                            <span>{item.label}</span>
                            {canEdit ? (
                              <div className="lookup-actions">
                                <button
                                  type="button"
                                  className="button secondary"
                                  onClick={() => openEditItem(item)}
                                >
                                  Edit
                                </button>
                                <button
                                  type="button"
                                  className="button secondary"
                                  onClick={() => deleteItem(item)}
                                >
                                  Delete
                                </button>
                              </div>
                            ) : null}
                          </div>
                        ))}
                      </div>
                    )}
                  </>
                ) : (
                  <p className="table-empty">No lookup fields configured.</p>
                )}
              </div>
            </div>
          )}

          {editor && mounted
            ? createPortal(
                <div className="dialog-backdrop" role="presentation">
                  <div
                    className="dialog-panel"
                    role="dialog"
                    aria-modal="true"
                    aria-labelledby="lookup-editor-title"
                  >
                    <div className="dialog-header">
                      <h3 id="lookup-editor-title" className="dialog-title">
                        {editor.mode === "add"
                          ? "Add sub-item"
                          : "Edit sub-item"}
                      </h3>
                      <button
                        type="button"
                        className="dialog-close"
                        aria-label="Close"
                        onClick={closeEditor}
                        disabled={saving}
                      >
                        ×
                      </button>
                    </div>

                    {error ? (
                      <p className="form-message error">{error}</p>
                    ) : null}

                    <form
                      className="client-form client-form--dialog"
                      onSubmit={submitEditor}
                    >
                      <label className="field">
                        <span className="field-label">
                          Label <span className="field-required">*</span>
                        </span>
                        <input
                          className="field-input"
                          value={editor.label}
                          onChange={(event) =>
                            setEditor({
                              ...editor,
                              label: event.target.value,
                            })
                          }
                          autoFocus
                          required
                        />
                      </label>

                      <div className="dialog-actions">
                        <button
                          type="button"
                          className="button secondary"
                          onClick={closeEditor}
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
        </>
      )}
    </section>
  );
}
