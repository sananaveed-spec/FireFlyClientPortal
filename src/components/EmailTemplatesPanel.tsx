"use client";

import { useMsal } from "@azure/msal-react";
import {
  useEffect,
  useMemo,
  useState,
  type FormEvent,
  type ReactNode,
} from "react";
import { useAccess } from "@/auth/accessContext";
import { SearchableChecklist } from "@/components/SearchableChecklist";
import {
  getActorFromAccount,
  logPortalAction,
} from "@/lib/history-client";
import {
  createEmptyEmailSignature,
  DEFAULT_SIGNATURE_WEBSITE,
  getSignatureLogoSrc,
  getSignatureWebsiteHref,
  hasEmailSignature,
  SIGNATURE_CONFIDENTIALITY_NOTICE,
  SIGNATURE_REGARDS_LINE,
  stripLegacySignatureClosing,
  type EmailSignature,
} from "@/lib/email-signatures";
import {
  createEmailTemplate,
  suggestNextTemplateLabel,
  type EmailTemplate,
  type EmailTemplatesStore,
} from "@/lib/email-templates";
import {
  getLookupOptions,
  type LookupDataStore,
} from "@/lib/lookup-data";
import { portalFetch } from "@/lib/portal-api";
import { PORTAL_USERS_CHANGED_EVENT } from "@/lib/portal-users";
import { formatMultiValue } from "@/lib/resource-clients";

/** Keep placeholders like [Period] visible as written in the template. */
function renderTemplateText(text: string): ReactNode {
  if (!text) {
    return null;
  }
  return text;
}

function EmailSignatureBlock({ signature }: { signature: EmailSignature }) {
  if (!hasEmailSignature(signature)) {
    return null;
  }
  const websiteHref = getSignatureWebsiteHref(signature.website);
  const logoSrc = getSignatureLogoSrc(signature);
  return (
    <div className="email-signature-block">
      {signature.name ? (
        <div className="email-signature-line email-signature-line--strong">
          {signature.name}
        </div>
      ) : null}
      {signature.email ? (
        <div className="email-signature-line email-signature-line--strong">
          {signature.email}
        </div>
      ) : null}
      {signature.position ? (
        <div className="email-signature-line email-signature-line--strong">
          {signature.position}
        </div>
      ) : null}
      {signature.phone ? (
        <div className="email-signature-line email-signature-line--strong">
          {signature.phone}
        </div>
      ) : null}
      {signature.website ? (
        <div className="email-signature-line">
          <a
            className="email-signature-link"
            href={websiteHref || DEFAULT_SIGNATURE_WEBSITE}
            target="_blank"
            rel="noreferrer"
          >
            {signature.website}
          </a>
        </div>
      ) : null}
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img
        className="email-signature-image"
        src={logoSrc}
        alt="FireFly Accountancy"
      />
      <p className="email-signature-disclaimer">
        {SIGNATURE_CONFIDENTIALITY_NOTICE}
      </p>
    </div>
  );
}

function BodyPreview({
  body,
  signature,
}: {
  body: string;
  signature: EmailSignature;
}) {
  const cleanedBody = stripLegacySignatureClosing(body);
  if (!cleanedBody.trim()) {
    return "—";
  }

  return (
    <>
      {renderTemplateText(cleanedBody)}
      {"\n\n"}
      <span>{SIGNATURE_REGARDS_LINE}</span>
      {"\n"}
      {hasEmailSignature(signature) ? (
        <EmailSignatureBlock signature={signature} />
      ) : (
        <p className="lookup-hint email-signature-hint">
          Set your email signature under My Profile to show it after Regards.
        </p>
      )}
    </>
  );
}

type DraftFields = {
  label: string;
  title: string;
  whenSent: string;
  subject: string;
  body: string;
  bookkeepingServices: string[];
};

type PanelMode = "view" | "edit" | "add";

function draftFromTemplate(template: EmailTemplate): DraftFields {
  return {
    label: template.label,
    title: template.title,
    whenSent: template.whenSent,
    subject: template.subject,
    body: stripLegacySignatureClosing(template.body),
    bookkeepingServices: [...(template.bookkeepingServices ?? [])],
  };
}

const EMPTY_DRAFT: DraftFields = {
  label: "",
  title: "",
  whenSent: "",
  subject: "",
  body: "",
  bookkeepingServices: [],
};

async function persistEmailTemplates(data: EmailTemplatesStore) {
  const response = await portalFetch("/api/email-templates", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  const payload = (await response.json()) as EmailTemplatesStore & {
    error?: string;
  };
  if (!response.ok) {
    throw new Error(payload.error || "Failed to save email templates.");
  }
  return payload;
}

export function EmailTemplatesPanel() {
  const { accounts } = useMsal();
  const { canEdit } = useAccess();
  const actor = getActorFromAccount(accounts[0]);
  const [store, setStore] = useState<EmailTemplatesStore>({ templates: [] });
  const [lookups, setLookups] = useState<LookupDataStore | null>(null);
  const [signature, setSignature] = useState<EmailSignature>(
    createEmptyEmailSignature(),
  );
  const [activeId, setActiveId] = useState<string>("template-1");
  const [mode, setMode] = useState<PanelMode>("view");
  const [draft, setDraft] = useState<DraftFields>(EMPTY_DRAFT);
  const [ready, setReady] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const bookkeepingOptions = useMemo(
    () => getLookupOptions(lookups, "bookkeeping"),
    [lookups],
  );

  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const [templatesRes, lookupsRes, signatureRes] = await Promise.all([
          fetch("/api/email-templates", { cache: "no-store" }),
          fetch("/api/lookup-data", { cache: "no-store" }),
          portalFetch("/api/email-signature", { cache: "no-store" }),
        ]);
        const data = (await templatesRes.json()) as EmailTemplatesStore & {
          error?: string;
        };
        const lookupsData = (await lookupsRes.json()) as LookupDataStore & {
          error?: string;
        };
        const signatureData = (await signatureRes.json()) as {
          signature?: EmailSignature;
          error?: string;
        };
        if (!templatesRes.ok) {
          throw new Error(data.error || "Failed to load email templates.");
        }
        if (!lookupsRes.ok) {
          throw new Error(lookupsData.error || "Failed to load lookup data.");
        }
        if (!cancelled) {
          const templates = data.templates ?? [];
          setStore({ templates });
          setLookups({ fields: lookupsData.fields ?? [] });
          if (signatureRes.ok && signatureData.signature) {
            setSignature(signatureData.signature);
          }
          if (templates[0]?.id) {
            setActiveId(templates[0].id);
          }
          setError(null);
        }
      } catch (loadError) {
        if (!cancelled) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load email templates.",
          );
        }
      } finally {
        if (!cancelled) {
          setReady(true);
        }
      }
    }

    async function reloadSignature() {
      try {
        const signatureRes = await portalFetch("/api/email-signature", {
          cache: "no-store",
        });
        const signatureData = (await signatureRes.json()) as {
          signature?: EmailSignature;
        };
        if (signatureRes.ok && signatureData.signature) {
          setSignature(signatureData.signature);
        }
      } catch {
        // Keep the existing signature if refresh fails.
      }
    }

    void load();

    function onUsersChanged() {
      void reloadSignature();
    }
    window.addEventListener(PORTAL_USERS_CHANGED_EVENT, onUsersChanged);

    return () => {
      cancelled = true;
      window.removeEventListener(PORTAL_USERS_CHANGED_EVENT, onUsersChanged);
    };
  }, []);

  const activeTemplate = useMemo(
    () =>
      store.templates.find((template) => template.id === activeId) ??
      store.templates[0] ??
      null,
    [store.templates, activeId],
  );

  const display =
    mode === "view" && activeTemplate
      ? draftFromTemplate(activeTemplate)
      : draft;

  const editing = mode === "edit" || mode === "add";

  async function saveStore(next: EmailTemplatesStore) {
    setSaving(true);
    setError(null);
    try {
      const saved = await persistEmailTemplates(next);
      setStore({ templates: saved.templates ?? [] });
      setMode("view");
    } catch (saveError) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save email templates.",
      );
    } finally {
      setSaving(false);
    }
  }

  function selectTemplate(id: string) {
    if (saving) {
      return;
    }
    setActiveId(id);
    setMode("view");
    setError(null);
  }

  function openAddTemplate() {
    setError(null);
    setDraft({
      ...EMPTY_DRAFT,
      label: suggestNextTemplateLabel(store.templates),
    });
    setMode("add");
  }

  function openEditTemplate(template: EmailTemplate) {
    setError(null);
    setDraft(draftFromTemplate(template));
    setMode("edit");
  }

  function cancelEditing() {
    if (saving) {
      return;
    }
    setError(null);
    setMode("view");
    if (activeTemplate) {
      setDraft(draftFromTemplate(activeTemplate));
    }
  }

  function deleteTemplate(template: EmailTemplate) {
    if (!window.confirm(`Delete "${template.label}: ${template.title}"?`)) {
      return;
    }
    const next = store.templates.filter((entry) => entry.id !== template.id);
    if (activeId === template.id) {
      setActiveId(next[0]?.id ?? "");
    }
    setMode("view");
    void saveStore({ templates: next });
    void logPortalAction({
      actor,
      project: "Email Templates",
      entityType: "lookupData",
      entityLabel: `${template.label} / ${template.title}`,
      action: "deleted",
      summary: `Email template deleted: ${template.title}`,
      details: `Label: ${template.label}\nTitle: ${template.title}`,
      comment: "Email template removed.",
    });
  }

  function updateDraft<K extends keyof DraftFields>(
    key: K,
    value: DraftFields[K],
  ) {
    setDraft((current) => ({ ...current, [key]: value }));
  }

  function submitForm(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    const label = draft.label.trim();
    const title = draft.title.trim();
    const whenSent = draft.whenSent.trim();
    const subject = draft.subject.trim();
    const body = stripLegacySignatureClosing(draft.body);
    const bookkeepingServices = draft.bookkeepingServices
      .map((item) => item.trim())
      .filter(Boolean);

    if (!label) {
      setError("Sidebar label is required.");
      return;
    }
    if (!title) {
      setError("Title is required.");
      return;
    }
    if (!subject) {
      setError("Subject is required.");
      return;
    }
    if (!body) {
      setError("Body is required.");
      return;
    }

    const payload = {
      label,
      title,
      whenSent,
      subject,
      body,
      bookkeepingServices,
    };

    if (mode === "add") {
      const created = createEmailTemplate(payload);
      void saveStore({ templates: [...store.templates, created] });
      setActiveId(created.id);
      void logPortalAction({
        actor,
        project: "Email Templates",
        entityType: "lookupData",
        entityLabel: `${label} / ${title}`,
        action: "created",
        summary: `Email template created: ${title}`,
        details: `Label: ${label}\nTitle: ${title}\nBookkeeping services: ${
          formatMultiValue(bookkeepingServices) || "—"
        }`,
        comment: "New email template added.",
      });
      return;
    }

    if (mode !== "edit" || !activeTemplate) {
      return;
    }

    const previous = activeTemplate.title;
    void saveStore({
      templates: store.templates.map((template) =>
        template.id === activeTemplate.id
          ? { ...template, ...payload }
          : template,
      ),
    });
    void logPortalAction({
      actor,
      project: "Email Templates",
      entityType: "lookupData",
      entityLabel: `${label} / ${title}`,
      action: "updated",
      summary: `Email template updated: ${title}`,
      details: `Label: ${label}\nTitle: "${previous || "—"}" → "${title}"\nBookkeeping services: ${
        formatMultiValue(bookkeepingServices) || "—"
      }`,
      comment: "Email template updated.",
    });
  }

  return (
    <>
      {error && !editing ? (
        <p className="form-message error">{error}</p>
      ) : null}
      {!ready ? (
        <p className="table-empty">Loading…</p>
      ) : (
        <div className="lookup-layout">
          <aside className="lookup-fields" aria-label="Email templates">
            {canEdit ? (
              <button
                type="button"
                className="lookup-field-link lookup-field-link--add"
                onClick={openAddTemplate}
                disabled={saving}
              >
                <span>Add New Template</span>
              </button>
            ) : null}
            {store.templates.map((template) => (
              <button
                key={template.id}
                type="button"
                className={
                  mode !== "add" && activeTemplate?.id === template.id
                    ? "lookup-field-link active"
                    : "lookup-field-link"
                }
                onClick={() => selectTemplate(template.id)}
              >
                <span>{template.label}</span>
              </button>
            ))}
          </aside>

          <div className="lookup-items email-template-panel">
            {mode === "add" || activeTemplate ? (
              <form className="email-template-form" onSubmit={submitForm}>
                <div className="email-template-header">
                  <div className="email-template-header__top">
                    <div>
                      <h3 className="section-title">
                        {mode === "add"
                          ? "Add New Template"
                          : `${display.label}: ${display.title || "Untitled"}`}
                      </h3>
                    </div>
                    {canEdit ? (
                      <div className="lookup-actions">
                        {editing ? (
                          <>
                            <button
                              type="button"
                              className="button secondary"
                              onClick={cancelEditing}
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
                          </>
                        ) : activeTemplate ? (
                          <>
                            <button
                              type="button"
                              className="button secondary"
                              onClick={() => openEditTemplate(activeTemplate)}
                              disabled={saving}
                            >
                              Edit
                            </button>
                            <button
                              type="button"
                              className="button secondary"
                              onClick={() => deleteTemplate(activeTemplate)}
                              disabled={saving}
                            >
                              Delete
                            </button>
                          </>
                        ) : null}
                      </div>
                    ) : null}
                  </div>
                </div>

                {error && editing ? (
                  <p className="form-message error">{error}</p>
                ) : null}

                <div className="email-template-fields">
                  {editing ? (
                    <>
                      <label className="field">
                        <span className="field-label">
                          Sidebar label{" "}
                          <span className="field-required">*</span>
                        </span>
                        <input
                          className="field-input"
                          value={draft.label}
                          onChange={(event) =>
                            updateDraft("label", event.target.value)
                          }
                          placeholder="Template 16"
                          required
                          disabled={saving}
                        />
                      </label>

                      <label className="field">
                        <span className="field-label">
                          Title <span className="field-required">*</span>
                        </span>
                        <input
                          className="field-input"
                          value={draft.title}
                          onChange={(event) =>
                            updateDraft("title", event.target.value)
                          }
                          required
                          disabled={saving}
                        />
                      </label>
                    </>
                  ) : null}

                  {editing ? (
                    <SearchableChecklist
                      label="Bookkeeping services"
                      values={draft.bookkeepingServices}
                      options={bookkeepingOptions}
                      placeholder="Search/Select from the list"
                      onChange={(values) =>
                        updateDraft("bookkeepingServices", values)
                      }
                    />
                  ) : (
                    <label className="field">
                      <span className="field-label">Bookkeeping services</span>
                      <div className="email-template-field-value">
                        {formatMultiValue(display.bookkeepingServices) || "—"}
                      </div>
                    </label>
                  )}

                  <label className="field">
                    <span className="field-label">When it is sent</span>
                    {editing ? (
                      <input
                        className="field-input"
                        value={draft.whenSent}
                        onChange={(event) =>
                          updateDraft("whenSent", event.target.value)
                        }
                        disabled={saving}
                      />
                    ) : (
                      <div className="email-template-field-value">
                        {display.whenSent || "—"}
                      </div>
                    )}
                  </label>

                  <label className="field">
                    <span className="field-label">Subject</span>
                    {editing ? (
                      <input
                        className="field-input"
                        value={draft.subject}
                        onChange={(event) =>
                          updateDraft("subject", event.target.value)
                        }
                        required
                        disabled={saving}
                      />
                    ) : (
                      <div className="email-template-field-value">
                        {renderTemplateText(display.subject) || "—"}
                      </div>
                    )}
                  </label>

                  <label className="field">
                    <span className="field-label">Body</span>
                    {editing ? (
                      <textarea
                        className="field-input email-template-body-input"
                        value={draft.body}
                        onChange={(event) =>
                          updateDraft("body", event.target.value)
                        }
                        rows={16}
                        required
                        disabled={saving}
                      />
                    ) : (
                      <div className="email-template-field-value email-template-field-value--body">
                        <BodyPreview
                          body={display.body}
                          signature={signature}
                        />
                      </div>
                    )}
                  </label>
                </div>
              </form>
            ) : (
              <p className="table-empty">
                No email templates yet
                {canEdit
                  ? '. Select "Add New Template" in the sidebar to create one.'
                  : "."}
              </p>
            )}
          </div>
        </div>
      )}
    </>
  );
}
