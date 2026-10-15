"use client";

import { useMsal } from "@azure/msal-react";
import {
  useCallback,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from "react";
import { mailSendRequest } from "@/auth/authConfig";
import { useAccess } from "@/auth/accessContext";
import { getAccountEmail } from "@/auth/organization";
import {
  formatBusinessDayLabel,
  getScheduleBusinessDate,
  scheduleBusinessDayNumber,
} from "@/lib/business-days";
import {
  ASSIGNMENT_ROLE_LABELS,
  getClientsAssignedToUser,
  staffNameAliases,
  type AssignedClient,
} from "@/lib/client-management";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import {
  addRequestBulletToList,
  completeActionItem,
  createEmptyRequestListsStore,
  findListContainingActionItem,
  getClientRequestList,
  normalizeClientRequestListsStore,
  removeRequestBullet,
  reopenActionItem,
  updateRequestBulletLabel,
  type ActionWorkflowItem,
  type ClientRequestList,
  type ClientRequestListsStore,
} from "@/lib/client-request-lists";
import {
  createEmptyEmailSignature,
  DEFAULT_SIGNATURE_WEBSITE,
  getSignatureLogoSrc,
  getSignatureWebsiteHref,
  hasEmailSignature,
  normalizeEmailSignature,
  SIGNATURE_CONFIDENTIALITY_NOTICE,
  stripLegacySignatureClosing,
  type EmailSignature,
} from "@/lib/email-signatures";
import {
  applyPortalLinkPlaceholders,
  PORTAL_WORD_MARKER,
  resolvePortalHref,
} from "@/lib/email-portal-link";
import {
  applyPeriodPlaceholder,
  createDefaultEmailTemplates,
  formatCurrentPeriod,
  normalizeEmailTemplates,
  type EmailTemplate,
  type EmailTemplatesStore,
} from "@/lib/email-templates";
import {
  createAdjustableSettingRow,
  createEmptyAdjustableSettingsStore,
  getClientAdjustableRows,
  normalizeFlexibleAdjustableSettingsStore,
  type AdjustableSettingRow,
  type FlexibleAdjustableSettingsStore,
} from "@/lib/flexible-adjustable-settings";
import {
  createDefaultHolidays,
  normalizeHolidays,
  type HolidaysStore,
} from "@/lib/holidays";
import { portalFetch } from "@/lib/portal-api";
import type { PortalUser } from "@/lib/portal-users";
import type { ResourceClientEntry } from "@/lib/resource-clients";
import {
  parseDateTimeLocalValue,
  sendGraphMail,
  toDateTimeLocalValue,
} from "@/lib/send-graph-mail";

/** Default Outlook send time: schedule business day at 8:00 local, or tomorrow 8:00 if past. */
function defaultScheduleSendAt(
  scheduleDate: Date | null,
  now: Date = new Date(),
): string {
  const candidate = new Date(scheduleDate ?? now);
  candidate.setHours(8, 0, 0, 0);
  if (candidate.getTime() <= now.getTime() + 60_000) {
    const tomorrow = new Date(now);
    tomorrow.setDate(tomorrow.getDate() + 1);
    tomorrow.setHours(8, 0, 0, 0);
    return toDateTimeLocalValue(tomorrow);
  }
  return toDateTimeLocalValue(candidate);
}

/** Bold auto-filled values (client name, period, etc.) inside subject/body. */
function renderAutofilledText(
  text: string,
  highlights: string[],
  portalHref?: string | null,
): ReactNode {
  if (!text) {
    return null;
  }

  function renderHighlighted(segment: string, keyPrefix: string): ReactNode {
    const unique = [
      ...new Set(highlights.map((value) => value.trim()).filter(Boolean)),
    ].sort((a, b) => b.length - a.length);
    if (unique.length === 0) {
      return segment;
    }
    const pattern = unique
      .map((value) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
      .join("|");
    const splitter = new RegExp(`(${pattern})`, "g");
    const isHighlight = new Set(unique);
    return segment.split(splitter).map((part, index) =>
      !part ? null : isHighlight.has(part) ? (
        <strong
          key={`${keyPrefix}-${part}-${index}`}
          className="email-template-autofill"
        >
          {part}
        </strong>
      ) : (
        <span key={`${keyPrefix}-${index}-${part.slice(0, 12)}`}>{part}</span>
      ),
    );
  }

  if (portalHref && text.includes(PORTAL_WORD_MARKER)) {
    const chunks = text.split(PORTAL_WORD_MARKER);
    const nodes: ReactNode[] = [];
    chunks.forEach((chunk, index) => {
      nodes.push(renderHighlighted(chunk, `chunk-${index}`));
      if (index < chunks.length - 1) {
        nodes.push(
          <a
            key={`portal-link-${index}`}
            className="email-signature-link"
            href={portalHref}
            target="_blank"
            rel="noreferrer"
          >
            portal
          </a>,
        );
      }
    });
    return nodes;
  }

  return renderHighlighted(text, "body");
}

function ClientScheduleEmailSignature({
  signature,
}: {
  signature: EmailSignature;
}) {
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

type WorkflowTab = "request" | "action" | "email" | "completed";
type ActionResolveTarget = "completed" | "late";

type ResourcesResponse = {
  clients?: ResourceClientEntry[];
  error?: string;
};

type ProfilesResponse = {
  profiles?: ClientProfileEntry[];
  error?: string;
};

type UsersResponse = {
  users?: PortalUser[];
  error?: string;
};

const CLIENT_INFO_FIELDS: { key: string; label: string }[] = [
  { key: "uniqueId", label: "Unique ID" },
  { key: "operatingName", label: "Operating Name" },
  { key: "industry", label: "Industry" },
  { key: "entityType", label: "Entity Type" },
  { key: "serviceFrequency", label: "Service Frequency" },
  { key: "reportingTier", label: "Reporting Tier" },
  { key: "paymentType", label: "Payment Type" },
  { key: "country", label: "Country" },
  { key: "reportingCurrency", label: "Reporting Currency" },
  { key: "state", label: "State/Province" },
  { key: "timeZone", label: "Client Time Zone" },
  { key: "yourTimeZone", label: "Your Time Zone" },
  { key: "businessAccountNumber", label: "Business Account Number" },
  { key: "role", label: "Your Role" },
  { key: "pointOfContact", label: "Point of Contact" },
  { key: "pocEmail", label: "POC Email" },
  { key: "portalLink", label: "Portal Link" },
];

const CLIENT_SCHEDULE_ITEMS = [
  { id: "bd1", label: "BD1" },
  { id: "bd5", label: "BD5" },
  { id: "bd7", label: "BD7" },
  { id: "bd8", label: "BD8" },
  { id: "questions-plus-3-bd", label: "Questions + 3 BD" },
  { id: "bd10", label: "BD10" },
  { id: "bd15", label: "BD15" },
  { id: "bd18", label: "BD18" },
  { id: "bd20", label: "BD20" },
  { id: "last-3-bd", label: "Last 3 BD of month" },
  { id: "last-2-bd", label: "Last 2 BD of month" },
  { id: "last-1-bd", label: "Last 1 BD of month" },
] as const;

const BD1_SUB_MENUS = [
  {
    id: "bank-credit-card",
    label: "Bank and Credit Card Statements Request",
  },
  {
    id: "other-statements",
    label:
      "Other Statements Request (Loans, Other Institutions, Leases)",
  },
] as const;

type Bd1SubMenuId = (typeof BD1_SUB_MENUS)[number]["id"];

function requestTypeScheduleId(
  scheduleId: string,
  subMenuId: Bd1SubMenuId,
): string {
  if (scheduleId === "bd1") {
    return subMenuId === "other-statements"
      ? "bd1-other-statements"
      : "bd1";
  }
  if (scheduleId === "bd5") {
    return subMenuId === "other-statements"
      ? "bd5-other-statements"
      : "bd5";
  }
  if (scheduleId === "bd7") {
    return subMenuId === "other-statements"
      ? "bd7-other-statements"
      : "bd7";
  }
  return scheduleId;
}

function bd1EmailTemplateId(subMenuId: Bd1SubMenuId): string {
  return subMenuId === "other-statements" ? "template-2" : "template-1";
}

function findBd1EmailTemplate(
  store: EmailTemplatesStore,
  subMenuId: Bd1SubMenuId,
): EmailTemplate | null {
  const templateId = bd1EmailTemplateId(subMenuId);
  return (
    store.templates.find((template) => template.id === templateId) ?? null
  );
}

function requestTypeForTemplateId(templateId: string): Bd1SubMenuId {
  return templateId === "template-2" ? "other-statements" : "bank-credit-card";
}

function findEmailTemplateById(
  store: EmailTemplatesStore,
  templateId: string,
): EmailTemplate | null {
  return (
    store.templates.find((template) => template.id === templateId) ?? null
  );
}

/** Flexible request lists are keyed by template so BD changes keep the same defaults. */
function flexibleScheduleIdForTemplate(templateId: string): string {
  return `flex-${templateId}`;
}

/** Client Management: BD1 send → BD5 Completed/Late review (= +4). */
const TEMPLATE_REVIEW_FOLLOW_UP_OFFSET = 4;
/** Client Management: BD1 send → BD7 Late Action Items (= +6). */
const TEMPLATE_LATE_FOLLOW_UP_OFFSET = 6;
/** Client Management: BD1 send → BD4 email due date (= +3). */
const TEMPLATE_DEADLINE_OFFSET = 3;

const TEMPLATES_WITH_FOLLOW_UP = new Set(["template-1", "template-2"]);

type TemplateFollowUpKind = "review" | "late";

function offsetBusinessDayId(
  scheduleId: string,
  offset: number,
): string | null {
  const n = scheduleBusinessDayNumber(scheduleId);
  if (n === null) {
    return null;
  }
  const next = n + offset;
  if (next < 1) {
    return null;
  }
  return `bd${next}`;
}

function businessDayLabelForId(scheduleId: string): string {
  const known = CLIENT_SCHEDULE_ITEMS.find((item) => item.id === scheduleId);
  if (known) {
    return known.label;
  }
  const n = scheduleBusinessDayNumber(scheduleId);
  return n !== null ? `BD${n}` : scheduleId.toUpperCase();
}

const TEMPLATE_1_DEFAULT_REQUEST_LABELS = [
  "Request List1",
  "Request List2",
  "Request List3",
  "Request List4",
  "Request List5",
] as const;

const TEMPLATE_2_DEFAULT_REQUEST_LABELS = [
  "List Action1",
  "List Action2",
  "List Action3",
] as const;

function toPortalHref(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }
  return `https://${trimmed}`;
}

function formatActionResolvedAt(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  const date = new Date(trimmed);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  return date.toLocaleString(undefined, {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

function clientInfoValue(
  selected: AssignedClient,
  key: string,
): ReactNode {
  const profile = selected.profile;
  const resource = selected.resource;

  switch (key) {
    case "operatingName":
      return profile?.operatingName?.trim() || "—";
    case "uniqueId":
      return profile?.uniqueId?.trim() || "—";
    case "businessAccountNumber":
      return profile?.businessAccountNumber?.trim() || "—";
    case "pointOfContact":
      return profile?.pointOfContact?.trim() || "—";
    case "pocEmail":
      return profile?.pocEmail?.trim() || "—";
    case "portalLink": {
      const link = profile?.portalLink?.trim() || "";
      if (!link) {
        return "—";
      }
      return (
        <a
          className="table-link"
          href={toPortalHref(link)}
          target="_blank"
          rel="noopener noreferrer"
        >
          {link}
        </a>
      );
    }
    case "industry":
      return profile?.industry?.trim() || resource.industry?.trim() || "—";
    case "entityType":
      return profile?.entityType?.trim() || "—";
    case "serviceFrequency":
      return profile?.serviceFrequency?.trim() || "—";
    case "reportingTier":
      return profile?.reportingTier?.trim() || "—";
    case "paymentType": {
      const paymentType = profile?.paymentType?.trim() || "";
      if (paymentType === "Cash") {
        return "Cash basis";
      }
      if (paymentType === "Accrual") {
        return "Accrual basis";
      }
      return paymentType || "—";
    }
    case "country":
      return profile?.country?.trim() || "—";
    case "reportingCurrency":
      return profile?.reportingCurrency?.trim() || "—";
    case "state":
      return profile?.state?.trim() || "—";
    case "timeZone":
      return profile?.timeZone?.trim() || "—";
    case "yourTimeZone":
      return "PKT";
    case "role":
      return selected.roles
        .map((role) => ASSIGNMENT_ROLE_LABELS[role])
        .join(" · ");
    default:
      return "—";
  }
}

function buildClientInfoRows(selected: AssignedClient) {
  const rows: Array<
    [
      { key: string; label: string; value: ReactNode },
      { key: string; label: string; value: ReactNode } | null,
    ]
  > = [];

  for (let index = 0; index < CLIENT_INFO_FIELDS.length; index += 2) {
    const left = CLIENT_INFO_FIELDS[index]!;
    const right = CLIENT_INFO_FIELDS[index + 1] ?? null;
    rows.push([
      {
        key: left.key,
        label: left.label,
        value: clientInfoValue(selected, left.key),
      },
      right
        ? {
            key: right.key,
            label: right.label,
            value: clientInfoValue(selected, right.key),
          }
        : null,
    ]);
  }

  return rows;
}

export function FlexibleClientManagementPanel() {
  const { instance, accounts } = useMsal();
  const { email, canEdit } = useAccess();
  const msalAccount = accounts[0] ?? null;
  const msalName = msalAccount?.name ?? "";
  const emailFrom =
    getAccountEmail(msalAccount) || email.trim().toLowerCase() || "";

  const [resources, setResources] = useState<ResourceClientEntry[]>([]);
  const [profiles, setProfiles] = useState<ClientProfileEntry[]>([]);
  const [portalName, setPortalName] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedScheduleId, setSelectedScheduleId] = useState<string>(
    CLIENT_SCHEDULE_ITEMS[0].id,
  );
  const [selectedSidebarRowId, setSelectedSidebarRowId] = useState<
    string | null
  >(null);
  const [requestLists, setRequestLists] = useState<ClientRequestListsStore>(
    createEmptyRequestListsStore(),
  );
  const [holidays, setHolidays] = useState<HolidaysStore>(createDefaultHolidays);
  const [emailTemplates, setEmailTemplates] = useState<EmailTemplatesStore>(
    createDefaultEmailTemplates,
  );
  const [emailSignature, setEmailSignature] = useState<EmailSignature>(
    createEmptyEmailSignature,
  );
  const [adjustableSettings, setAdjustableSettings] =
    useState<FlexibleAdjustableSettingsStore>(
      createEmptyAdjustableSettingsStore,
    );
  const [adjustableDraft, setAdjustableDraft] = useState<{
    id: string | null;
    templateId: string;
    businessDayId: string;
  } | null>(null);
  const [selectedAdjustableRowIds, setSelectedAdjustableRowIds] = useState<
    string[]
  >([]);
  const [savingAdjustable, setSavingAdjustable] = useState(false);
  const [adjustableMessage, setAdjustableMessage] = useState<string | null>(
    null,
  );
  const [adjustableError, setAdjustableError] = useState<string | null>(null);
  const [workflowTab, setWorkflowTab] = useState<WorkflowTab>("action");
  const [requestTypeId, setRequestTypeId] =
    useState<Bd1SubMenuId>("bank-credit-card");
  const [draftRequestLabel, setDraftRequestLabel] = useState("");
  const [editingRequestItemId, setEditingRequestItemId] = useState<
    string | null
  >(null);
  const [editRequestLabel, setEditRequestLabel] = useState("");
  const [savingRequestList, setSavingRequestList] = useState(false);
  const [emailTo, setEmailTo] = useState("");
  const [emailCc, setEmailCc] = useState("");
  const [emailBcc, setEmailBcc] = useState("");
  const [emailDeadline, setEmailDeadline] = useState("");
  const [emailSendAt, setEmailSendAt] = useState("");
  const [emailSubject, setEmailSubject] = useState("");
  const [emailBody, setEmailBody] = useState("");
  /** Period month folder URL (existing SharePoint year/month path). */
  const [emailPortalHref, setEmailPortalHref] = useState<string | null>(null);
  const [emailAutofillHighlights, setEmailAutofillHighlights] = useState<
    string[]
  >([]);
  const [clientInfoOpen, setClientInfoOpen] = useState(false);
  const [adjustableSettingsOpen, setAdjustableSettingsOpen] = useState(false);
  const [sendingEmail, setSendingEmail] = useState(false);
  const [schedulingEmail, setSchedulingEmail] = useState(false);
  const [emailSendMessage, setEmailSendMessage] = useState<string | null>(null);
  const [emailSendError, setEmailSendError] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setEditingRequestItemId(null);
    setEditRequestLabel("");
  }, [selectedId, selectedScheduleId, workflowTab, requestTypeId]);

  const load = useCallback(async () => {
    setError(null);
    try {
      const [
        resourcesRes,
        profilesRes,
        usersRes,
        requestListsRes,
        holidaysRes,
        emailTemplatesRes,
        signatureRes,
        adjustableRes,
      ] = await Promise.all([
        fetch("/api/resource-clients", { cache: "no-store" }),
        fetch("/api/client-profiles", { cache: "no-store" }),
        fetch("/api/portal-users", { cache: "no-store" }),
        fetch("/api/client-request-lists", { cache: "no-store" }),
        fetch("/api/holidays", { cache: "no-store" }),
        fetch("/api/email-templates", { cache: "no-store" }),
        portalFetch("/api/email-signature", { cache: "no-store" }),
        fetch("/api/flexible-adjustable-settings", { cache: "no-store" }),
      ]);
      const resourcesData = (await resourcesRes.json()) as ResourcesResponse;
      const profilesData = (await profilesRes.json()) as ProfilesResponse;
      const usersData = (await usersRes.json()) as UsersResponse;
      const requestListsData = await requestListsRes.json();
      const holidaysData = await holidaysRes.json();
      const emailTemplatesData = await emailTemplatesRes.json();
      const adjustableData = await adjustableRes.json();
      const signatureData = (await signatureRes.json()) as {
        signature?: EmailSignature;
        error?: string;
      };

      if (!resourcesRes.ok) {
        throw new Error(
          resourcesData.error || "Failed to load resource allocation.",
        );
      }
      if (!profilesRes.ok) {
        throw new Error(profilesData.error || "Failed to load client profiles.");
      }

      const nextResources = Array.isArray(resourcesData.clients)
        ? resourcesData.clients
        : [];
      const nextProfiles = Array.isArray(profilesData.profiles)
        ? profilesData.profiles
        : [];
      const matchedUser =
        usersData.users?.find((user) => user.email === email) ?? null;

      setResources(nextResources);
      setProfiles(nextProfiles);
      setPortalName(matchedUser?.name?.trim() || "");
      setRequestLists(
        requestListsRes.ok
          ? normalizeClientRequestListsStore(requestListsData)
          : createEmptyRequestListsStore(),
      );
      setHolidays(
        holidaysRes.ok
          ? normalizeHolidays(holidaysData)
          : createDefaultHolidays(),
      );
      setEmailTemplates(
        emailTemplatesRes.ok
          ? normalizeEmailTemplates(emailTemplatesData)
          : createDefaultEmailTemplates(),
      );
      setEmailSignature(
        signatureRes.ok
          ? normalizeEmailSignature(signatureData.signature)
          : createEmptyEmailSignature(),
      );
      setAdjustableSettings(
        adjustableRes.ok
          ? normalizeFlexibleAdjustableSettingsStore(adjustableData)
          : createEmptyAdjustableSettingsStore(),
      );
      setSelectedId((current) =>
        current && nextResources.some((entry) => entry.id === current)
          ? current
          : null,
      );
    } catch (loadError) {
      setError(
        loadError instanceof Error
          ? loadError.message
          : "Failed to load client management.",
      );
    } finally {
      setReady(true);
    }
  }, [email]);

  useEffect(() => {
    void load();
  }, [load]);

  const aliases = useMemo(
    () => staffNameAliases(portalName, msalName),
    [portalName, msalName],
  );

  const assignedClients = useMemo(
    () => getClientsAssignedToUser(resources, profiles, aliases),
    [resources, profiles, aliases],
  );

  useEffect(() => {
    if (assignedClients.length === 0) {
      setSelectedId(null);
      return;
    }
    const stillValid =
      selectedId &&
      assignedClients.some((entry) => entry.resource.id === selectedId);
    if (!stillValid) {
      setSelectedId(assignedClients[0]?.resource.id ?? null);
    }
  }, [assignedClients, selectedId]);

  const selected: AssignedClient | null =
    assignedClients.find((entry) => entry.resource.id === selectedId) ??
    assignedClients[0] ??
    null;

  const selectedClientId = selected?.resource.id ?? "";
  const selectedOperatingName =
    selected?.profile?.operatingName?.trim() ||
    selected?.resource.clientName?.trim() ||
    "";
  const selectedPointOfContact =
    selected?.profile?.pointOfContact?.trim() || "";
  const selectedPortalLink = selected?.profile?.portalLink?.trim() || "";

  const adjustableRows = useMemo(
    () => getClientAdjustableRows(adjustableSettings, selectedClientId),
    [adjustableSettings, selectedClientId],
  );

  useEffect(() => {
    setAdjustableDraft(null);
    setSelectedAdjustableRowIds([]);
    setSelectedSidebarRowId(null);
    setAdjustableMessage(null);
    setAdjustableError(null);
  }, [selectedClientId]);

  useEffect(() => {
    setSelectedAdjustableRowIds((current) =>
      current.filter((id) => adjustableRows.some((row) => row.id === id)),
    );
  }, [adjustableRows]);

  const selectedAdjustableRows = useMemo(
    () =>
      adjustableRows.filter((row) => selectedAdjustableRowIds.includes(row.id)),
    [adjustableRows, selectedAdjustableRowIds],
  );
  const selectedAdjustableRow =
    selectedAdjustableRows.length === 1
      ? selectedAdjustableRows[0] ?? null
      : null;
  const allAdjustableSelected =
    adjustableRows.length > 0 &&
    selectedAdjustableRowIds.length === adjustableRows.length;

  function toggleAdjustableRowSelected(rowId: string) {
    setSelectedAdjustableRowIds((current) =>
      current.includes(rowId)
        ? current.filter((id) => id !== rowId)
        : [...current, rowId],
    );
  }

  function toggleSelectAllAdjustableRows() {
    setSelectedAdjustableRowIds(
      allAdjustableSelected ? [] : adjustableRows.map((row) => row.id),
    );
  }

  /** Adjustable Setting rows → sidebar BDs + optional multi-template chooser. */
  const sidebarTemplateItems = useMemo(() => {
    return adjustableRows.map((row) => {
      const template = findEmailTemplateById(emailTemplates, row.templateId);
      const schedule = CLIENT_SCHEDULE_ITEMS.find(
        (item) => item.id === row.businessDayId,
      );
      return {
        rowId: row.id,
        templateId: row.templateId,
        businessDayId: row.businessDayId,
        label: template
          ? template.title.trim() || template.label
          : row.templateId,
        businessDayLabel: schedule?.label ?? row.businessDayId,
      };
    });
  }, [adjustableRows, emailTemplates]);

  const selectedSidebarItem =
    sidebarTemplateItems.find((item) => item.rowId === selectedSidebarRowId) ??
    sidebarTemplateItems[0] ??
    null;

  const selectedTemplateId = selectedSidebarItem?.templateId ?? "";
  const isTemplate1 = selectedTemplateId === "template-1";
  const isTemplate2 = selectedTemplateId === "template-2";

  /** Template 1/2 on BDn → BDn+4 review and BDn+6 late follow-ups. */
  const templateFollowUps = useMemo(() => {
    const out: {
      sourceRowId: string;
      templateId: string;
      sendBusinessDayId: string;
      followUpBusinessDayId: string;
      kind: TemplateFollowUpKind;
      label: string;
    }[] = [];
    for (const item of sidebarTemplateItems) {
      if (!TEMPLATES_WITH_FOLLOW_UP.has(item.templateId)) {
        continue;
      }
      const offsets: { kind: TemplateFollowUpKind; offset: number }[] = [
        { kind: "review", offset: TEMPLATE_REVIEW_FOLLOW_UP_OFFSET },
        { kind: "late", offset: TEMPLATE_LATE_FOLLOW_UP_OFFSET },
      ];
      for (const { kind, offset } of offsets) {
        const followUpBusinessDayId = offsetBusinessDayId(
          item.businessDayId,
          offset,
        );
        if (!followUpBusinessDayId) {
          continue;
        }
        out.push({
          sourceRowId: item.rowId,
          templateId: item.templateId,
          sendBusinessDayId: item.businessDayId,
          followUpBusinessDayId,
          kind,
          label: item.label,
        });
      }
    }
    return out;
  }, [sidebarTemplateItems]);

  const activeTemplateFollowUp =
    (isTemplate1 || isTemplate2) && selectedSidebarItem
      ? (templateFollowUps.find(
          (entry) =>
            entry.sourceRowId === selectedSidebarItem.rowId &&
            entry.followUpBusinessDayId === selectedScheduleId,
        ) ?? null)
      : null;
  const isTemplateReviewFollowUpDay =
    activeTemplateFollowUp?.kind === "review";
  const isTemplateLateFollowUpDay = activeTemplateFollowUp?.kind === "late";
  const isTemplateFollowUpDay =
    isTemplateReviewFollowUpDay || isTemplateLateFollowUpDay;

  const showBd5Content =
    isTemplateReviewFollowUpDay ||
    (selectedScheduleId === "bd5" && !isTemplate1 && !isTemplate2);
  const showBd7Content =
    isTemplateLateFollowUpDay ||
    (selectedScheduleId === "bd7" && !isTemplate1 && !isTemplate2);
  const showBd1LikeWorkflow =
    ((isTemplate1 || isTemplate2) && !isTemplateReviewFollowUpDay) ||
    showBd7Content;
  const templatesForSelectedBd = useMemo(() => {
    const onBd = sidebarTemplateItems.filter(
      (item) => item.businessDayId === selectedScheduleId,
    );
    const merged = [...onBd];
    const seenRowIds = new Set(merged.map((item) => item.rowId));
    for (const followUp of templateFollowUps) {
      if (followUp.followUpBusinessDayId !== selectedScheduleId) {
        continue;
      }
      if (seenRowIds.has(followUp.sourceRowId)) {
        continue;
      }
      const source = sidebarTemplateItems.find(
        (item) => item.rowId === followUp.sourceRowId,
      );
      if (source) {
        seenRowIds.add(source.rowId);
        merged.push(source);
      }
    }
    return merged;
  }, [sidebarTemplateItems, selectedScheduleId, templateFollowUps]);
  // Multiple templates on the same BD → Step 1 cards (like request-type chooser).
  const showTemplateChooser = templatesForSelectedBd.length > 1;
  const showRequestListTab =
    (isTemplate1 || isTemplate2) && !isTemplateFollowUpDay;
  const activeRequestType =
    BD1_SUB_MENUS.find((item) => item.id === requestTypeId) ??
    BD1_SUB_MENUS[0];
  const activeRequestScheduleId = selectedTemplateId
    ? flexibleScheduleIdForTemplate(selectedTemplateId)
    : selectedScheduleId;

  const sidebarScheduleItems = useMemo(() => {
    const seen = new Set<string>();
    const items: { id: string; label: string }[] = [];
    for (const item of sidebarTemplateItems) {
      if (seen.has(item.businessDayId)) {
        continue;
      }
      seen.add(item.businessDayId);
      items.push({
        id: item.businessDayId,
        label: item.businessDayLabel,
      });
    }
    for (const followUp of templateFollowUps) {
      if (seen.has(followUp.followUpBusinessDayId)) {
        continue;
      }
      seen.add(followUp.followUpBusinessDayId);
      items.push({
        id: followUp.followUpBusinessDayId,
        label: businessDayLabelForId(followUp.followUpBusinessDayId),
      });
    }
    return items.sort((a, b) => {
      const aN = scheduleBusinessDayNumber(a.id);
      const bN = scheduleBusinessDayNumber(b.id);
      if (aN !== null && bN !== null) {
        return aN - bN;
      }
      if (aN !== null) {
        return -1;
      }
      if (bN !== null) {
        return 1;
      }
      return a.label.localeCompare(b.label);
    });
  }, [sidebarTemplateItems, templateFollowUps]);

  useEffect(() => {
    if (sidebarTemplateItems.length === 0) {
      setSelectedSidebarRowId(null);
      return;
    }
    const scheduleIds = sidebarScheduleItems.map((item) => item.id);
    const scheduleId = scheduleIds.includes(selectedScheduleId)
      ? selectedScheduleId
      : scheduleIds[0]!;
    if (scheduleId !== selectedScheduleId) {
      setSelectedScheduleId(scheduleId);
      return;
    }
    const onBd = sidebarTemplateItems.filter(
      (item) => item.businessDayId === scheduleId,
    );
    const followUpSourceIds = new Set<string>();
    const followUpSources: (typeof sidebarTemplateItems)[number][] = [];
    for (const entry of templateFollowUps) {
      if (entry.followUpBusinessDayId !== scheduleId) {
        continue;
      }
      if (followUpSourceIds.has(entry.sourceRowId)) {
        continue;
      }
      const source = sidebarTemplateItems.find(
        (item) => item.rowId === entry.sourceRowId,
      );
      if (!source) {
        continue;
      }
      followUpSourceIds.add(entry.sourceRowId);
      followUpSources.push(source);
    }
    const choices = [
      ...onBd,
      ...followUpSources.filter(
        (source) => !onBd.some((item) => item.rowId === source.rowId),
      ),
    ];
    if (choices.length === 0) {
      return;
    }
    const active =
      choices.find((item) => item.rowId === selectedSidebarRowId) ??
      choices[0]!;
    if (selectedSidebarRowId !== active.rowId) {
      setSelectedSidebarRowId(active.rowId);
    }
    setRequestTypeId(requestTypeForTemplateId(active.templateId));
  }, [
    sidebarTemplateItems,
    sidebarScheduleItems,
    templateFollowUps,
    selectedScheduleId,
    selectedSidebarRowId,
  ]);

  useEffect(() => {
    if (!showRequestListTab && workflowTab === "request") {
      setWorkflowTab("action");
    }
  }, [showRequestListTab, workflowTab]);

  useEffect(() => {
    if (workflowTab === "email" && !showBd1LikeWorkflow) {
      setWorkflowTab("action");
    }
  }, [workflowTab, showBd1LikeWorkflow]);

  useEffect(() => {
    if (showBd1LikeWorkflow && workflowTab === "completed") {
      setWorkflowTab("action");
    }
  }, [showBd1LikeWorkflow, workflowTab]);

  useEffect(() => {
    if (showBd5Content && workflowTab === "completed") {
      setWorkflowTab("action");
    }
  }, [showBd5Content, workflowTab]);

  useEffect(() => {
    if (isTemplateReviewFollowUpDay && workflowTab === "email") {
      setWorkflowTab("action");
    }
  }, [isTemplateReviewFollowUpDay, workflowTab]);

  const defaultEmailDeadline = useMemo(() => {
    if (
      !selectedClientId ||
      isTemplateReviewFollowUpDay ||
      (!isTemplate1 && !isTemplate2 && !showBd7Content)
    ) {
      return "";
    }
    const deadlineScheduleId =
      isTemplate1 || isTemplate2
        ? offsetBusinessDayId(
            selectedSidebarItem?.businessDayId ?? selectedScheduleId,
            TEMPLATE_DEADLINE_OFFSET,
          ) ?? "bd4"
        : "bd4";
    const deadlineDate = getScheduleBusinessDate(
      deadlineScheduleId,
      holidays,
      selected?.profile?.country,
      selected?.profile?.state ?? "",
    );
    if (!deadlineDate) {
      return "";
    }
    const year = deadlineDate.getFullYear();
    const month = String(deadlineDate.getMonth() + 1).padStart(2, "0");
    const day = String(deadlineDate.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }, [
    selectedClientId,
    selectedScheduleId,
    selectedSidebarItem?.businessDayId,
    isTemplate1,
    isTemplate2,
    isTemplateReviewFollowUpDay,
    showBd7Content,
    selected?.profile?.country,
    selected?.profile?.state,
    holidays,
  ]);

  const effectiveEmailDeadline = emailDeadline || defaultEmailDeadline;

  const emailDeadlineLabel = useMemo(() => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(effectiveEmailDeadline)) {
      return "[Due Date]";
    }
    const [year, month, day] = effectiveEmailDeadline.split("-").map(Number);
    return formatBusinessDayLabel(new Date(year, month - 1, day));
  }, [effectiveEmailDeadline]);

  useEffect(() => {
    if (
      !selectedClientId ||
      isTemplateReviewFollowUpDay ||
      (!isTemplate1 && !isTemplate2 && !showBd7Content)
    ) {
      return;
    }
    const template =
      (selectedSidebarItem
        ? findEmailTemplateById(emailTemplates, selectedSidebarItem.templateId)
        : null) ?? findBd1EmailTemplate(emailTemplates, requestTypeId);
    const clientName =
      selectedOperatingName.replace(/[.:]+$/g, "").trim() ||
      selectedOperatingName.trim() ||
      "[Client Name]";
    const contactName = selectedPointOfContact || "[Contact First Name]";
    const portalLink = selectedPortalLink || "[Portal Link]";
    const periodDate = new Date();
    const period = formatCurrentPeriod(periodDate);
    const periodPortalHref = resolvePortalHref(portalLink, periodDate);
    const dueDate = emailDeadlineLabel;
    const flexList = getClientRequestList(
      requestLists,
      selectedClientId,
      selectedTemplateId
        ? flexibleScheduleIdForTemplate(selectedTemplateId)
        : requestTypeScheduleId("bd1", requestTypeId),
    );
    const emailListItems = showBd7Content
      ? flexList.actionItems.filter((item) => item.status === "late")
      : flexList.actionItems.filter((item) => item.status === "action");
    const bankListPlaceholder =
      "• [List of bank and credit card accounts, with last 4 digits]";
    const otherListPlaceholder =
      "• [List of loan, line of credit, lease, investment or other accounts]";
    const actionListText =
      emailListItems.length > 0
        ? emailListItems.map((item) => `• ${item.label}`).join("\n")
        : requestTypeId === "other-statements"
          ? otherListPlaceholder
          : bankListPlaceholder;

    function fillPlaceholders(text: string) {
      return applyPeriodPlaceholder(
        applyPortalLinkPlaceholders(
          text
            .replace(/\[Client Name\]/gi, clientName)
            .replace(/\[Contact First Name\]/gi, contactName)
            .replace(/\[Due Date\]/gi, dueDate)
            .replace(/\[Deadline\]/gi, dueDate)
            .replace(
              /(?:•\s*)?\[List of bank and credit card accounts, with last 4 digits\]/gi,
              actionListText,
            )
            .replace(
              /(?:•\s*)?\[List of loan, line of credit, lease, investment or other accounts\]/gi,
              actionListText,
            ),
          portalLink,
          periodDate,
        ),
      );
    }

    const subject = fillPlaceholders(template?.subject?.trim() || "");
    const body = fillPlaceholders(
      stripLegacySignatureClosing(template?.body ?? "").trimEnd(),
    );
    const highlights = [
      clientName,
      contactName,
      period,
      dueDate,
      ...emailListItems.map((item) => item.label),
    ].filter(
      (value) =>
        value &&
        value !== "[Client Name]" &&
        value !== "[Contact First Name]" &&
        value !== "[Due Date]" &&
        value !== "[Deadline]" &&
        value !== "[Portal Link]",
    );

    setEmailSubject(subject);
    setEmailBody(body);
    setEmailPortalHref(periodPortalHref);
    setEmailAutofillHighlights(highlights);
  }, [
    selectedClientId,
    selectedScheduleId,
    requestTypeId,
    isTemplate1,
    isTemplate2,
    isTemplateReviewFollowUpDay,
    showBd7Content,
    selectedSidebarItem,
    selectedOperatingName,
    selectedPointOfContact,
    selectedPortalLink,
    emailDeadlineLabel,
    requestLists,
    emailTemplates,
    selectedTemplateId,
  ]);

  useEffect(() => {
    if (
      !selectedClientId ||
      isTemplateReviewFollowUpDay ||
      (!isTemplate1 && !isTemplate2 && !showBd7Content)
    ) {
      return;
    }
    // Leave To empty on load; From always follows the signed-in account.
    setEmailTo("");
    setEmailCc("");
    setEmailBcc("");
    setEmailDeadline("");
    setEmailSendAt("");
    setEmailSendError(null);
    setEmailSendMessage(null);
  }, [
    selectedClientId,
    selectedScheduleId,
    requestTypeId,
    isTemplate1,
    isTemplate2,
    isTemplateReviewFollowUpDay,
    showBd7Content,
    selectedTemplateId,
  ]);

  const selectedSchedule =
    sidebarScheduleItems.find((item) => item.id === selectedScheduleId) ??
    sidebarScheduleItems[0] ??
    CLIENT_SCHEDULE_ITEMS.find((item) => item.id === selectedScheduleId) ??
    (selectedScheduleId
      ? {
          id: selectedScheduleId,
          label: businessDayLabelForId(selectedScheduleId),
        }
      : CLIENT_SCHEDULE_ITEMS[0]);

  const selectedScheduleBusinessDate = useMemo(() => {
    if (!selected || !selectedScheduleId || sidebarScheduleItems.length === 0) {
      return null;
    }
    return getScheduleBusinessDate(
      selectedScheduleId,
      holidays,
      selected.profile?.country,
      selected.profile?.state ?? "",
    );
  }, [
    selected,
    selectedScheduleId,
    holidays,
    sidebarScheduleItems.length,
  ]);

  useEffect(() => {
    if (!showBd1LikeWorkflow || isTemplateReviewFollowUpDay) {
      return;
    }
    setEmailSendAt(defaultScheduleSendAt(selectedScheduleBusinessDate));
  }, [
    selectedClientId,
    selectedScheduleId,
    selectedScheduleBusinessDate,
    showBd1LikeWorkflow,
    isTemplateReviewFollowUpDay,
  ]);

  const scheduleDateById = useMemo(() => {
    const map = new Map<string, string>();
    if (!selected) {
      return map;
    }
    const country = selected.profile?.country;
    const state = selected.profile?.state ?? "";
    for (const item of sidebarScheduleItems) {
      const date = getScheduleBusinessDate(
        item.id,
        holidays,
        country,
        state,
      );
      if (date) {
        map.set(item.id, formatBusinessDayLabel(date));
      }
    }
    return map;
  }, [selected, holidays, sidebarScheduleItems]);

  const selectedScheduleDate =
    selectedSchedule
      ? scheduleDateById.get(selectedSchedule.id) ?? null
      : null;

  const activeList: ClientRequestList =
    selected && activeRequestScheduleId
      ? getClientRequestList(
          requestLists,
          selected.resource.id,
          activeRequestScheduleId,
        )
      : {
          clientId: "",
          scheduleId: "",
          requestItems: [],
          actionItems: [],
          lastPromotedMonth: "",
        };

  const scopedActionItems = activeList.actionItems;
  const actionItems = scopedActionItems.filter(
    (item) => item.status === "action",
  );
  const completedItems = scopedActionItems.filter(
    (item) => item.status === "completed",
  );
  const bd7LateItems = useMemo(() => {
    return scopedActionItems
      .filter((item) => item.status === "late")
      .sort((a, b) => {
        const aTime = Date.parse(a.completedAt) || 0;
        const bTime = Date.parse(b.completedAt) || 0;
        return bTime - aTime;
      });
  }, [scopedActionItems]);
  const bd5ActionTableItems = useMemo(() => {
    const open = scopedActionItems.filter((item) => item.status === "action");
    const resolved = scopedActionItems
      .filter((item) => item.status === "completed" || item.status === "late")
      .sort((a, b) => {
        const aTime = Date.parse(a.completedAt) || 0;
        const bTime = Date.parse(b.completedAt) || 0;
        return bTime - aTime;
      });
    return [...open, ...resolved];
  }, [scopedActionItems]);

  async function persistList(nextList: ClientRequestList) {
    if (!selected || !activeRequestScheduleId) {
      return;
    }
    setSavingRequestList(true);
    setError(null);
    try {
      const response = await portalFetch("/api/client-request-lists", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(nextList),
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(
          (data as { error?: string }).error || "Failed to save request list.",
        );
      }
      setRequestLists(normalizeClientRequestListsStore(data));
    } catch (saveError) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save request list.",
      );
    } finally {
      setSavingRequestList(false);
    }
  }

  // Template 1 / 2 defaults stick to the template, not the business day.
  useEffect(() => {
    if (
      !selected ||
      !activeRequestScheduleId ||
      savingRequestList ||
      (!isTemplate1 && !isTemplate2)
    ) {
      return;
    }
    const alreadyStored = requestLists.lists.some(
      (list) =>
        list.clientId === selected.resource.id &&
        list.scheduleId === activeRequestScheduleId,
    );
    if (alreadyStored) {
      return;
    }

    const defaults = isTemplate1
      ? TEMPLATE_1_DEFAULT_REQUEST_LABELS
      : TEMPLATE_2_DEFAULT_REQUEST_LABELS;

    let next: ClientRequestList = {
      clientId: selected.resource.id,
      scheduleId: activeRequestScheduleId,
      requestItems: [],
      actionItems: [],
      lastPromotedMonth: "",
    };
    for (const label of defaults) {
      next = addRequestBulletToList(next, label);
    }
    void persistList(next);
  }, [
    selected,
    isTemplate1,
    isTemplate2,
    activeRequestScheduleId,
    requestLists,
    savingRequestList,
  ]);

  function templateOptionLabel(template: EmailTemplate) {
    const title = template.title.trim();
    return title ? `${template.label} — ${title}` : template.label;
  }

  function businessDayLabel(businessDayId: string) {
    return (
      CLIENT_SCHEDULE_ITEMS.find((item) => item.id === businessDayId)?.label ??
      businessDayId
    );
  }

  function templateLabelForId(templateId: string) {
    const template = emailTemplates.templates.find(
      (item) => item.id === templateId,
    );
    return template ? templateOptionLabel(template) : templateId;
  }

  async function persistAdjustableRows(nextRows: AdjustableSettingRow[]) {
    if (!selectedClientId) {
      return;
    }
    setSavingAdjustable(true);
    setAdjustableError(null);
    setAdjustableMessage(null);
    try {
      const response = await portalFetch("/api/flexible-adjustable-settings", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          clientId: selectedClientId,
          rows: nextRows,
        }),
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(
          (data as { error?: string }).error ||
            "Failed to save adjustable settings.",
        );
      }
      setAdjustableSettings(normalizeFlexibleAdjustableSettingsStore(data));
      setAdjustableDraft(null);
      setAdjustableMessage("Adjustable settings saved.");
    } catch (saveError) {
      setAdjustableError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save adjustable settings.",
      );
    } finally {
      setSavingAdjustable(false);
    }
  }

  function startAddAdjustableRow() {
    setAdjustableError(null);
    setAdjustableMessage(null);
    setSelectedAdjustableRowIds([]);
    setAdjustableDraft({
      id: null,
      templateId: emailTemplates.templates[0]?.id ?? "",
      businessDayId: CLIENT_SCHEDULE_ITEMS[0]?.id ?? "bd1",
    });
  }

  function startEditSelectedAdjustableRow() {
    if (!selectedAdjustableRow) {
      setAdjustableError("Select exactly one row to edit.");
      return;
    }
    setAdjustableError(null);
    setAdjustableMessage(null);
    setAdjustableDraft({
      id: selectedAdjustableRow.id,
      templateId: selectedAdjustableRow.templateId,
      businessDayId: selectedAdjustableRow.businessDayId,
    });
  }

  async function saveAdjustableDraft() {
    if (!adjustableDraft) {
      return;
    }
    if (
      !adjustableDraft.templateId.trim() ||
      !adjustableDraft.businessDayId.trim()
    ) {
      setAdjustableError("Select a template and a business day.");
      return;
    }

    if (adjustableDraft.id) {
      const nextRows = adjustableRows.map((row) =>
        row.id === adjustableDraft.id
          ? {
              ...row,
              templateId: adjustableDraft.templateId,
              businessDayId: adjustableDraft.businessDayId,
            }
          : row,
      );
      await persistAdjustableRows(nextRows);
      setSelectedAdjustableRowIds([adjustableDraft.id]);
      return;
    }

    const created = createAdjustableSettingRow({
      templateId: adjustableDraft.templateId,
      businessDayId: adjustableDraft.businessDayId,
    });
    if (!created) {
      setAdjustableError("Select a template and a business day.");
      return;
    }
    await persistAdjustableRows([...adjustableRows, created]);
    setSelectedAdjustableRowIds([created.id]);
  }

  async function deleteSelectedAdjustableRows() {
    if (selectedAdjustableRowIds.length === 0) {
      setAdjustableError("Select one or more rows to delete.");
      return;
    }
    const count = selectedAdjustableRowIds.length;
    const confirmLabel =
      count === 1
        ? "Delete this adjustable setting?"
        : `Delete ${count} adjustable settings?`;
    if (!window.confirm(confirmLabel)) {
      return;
    }
    const removing = new Set(selectedAdjustableRowIds);
    await persistAdjustableRows(
      adjustableRows.filter((row) => !removing.has(row.id)),
    );
    setSelectedAdjustableRowIds([]);
  }

  async function handleAddRequestItem() {
    const label = draftRequestLabel.trim();
    if (!label || !selected || !activeRequestScheduleId) {
      return;
    }
    const nextList = addRequestBulletToList(
      {
        ...activeList,
        clientId: selected.resource.id,
        scheduleId: activeRequestScheduleId,
      },
      label,
    );
    setDraftRequestLabel("");
    await persistList(nextList);
  }

  async function handleMoveRequestItem(itemId: string, direction: -1 | 1) {
    if (!selected || !activeRequestScheduleId) {
      return;
    }
    const items = [...activeList.requestItems];
    const index = items.findIndex((item) => item.id === itemId);
    const target = index + direction;
    if (index < 0 || target < 0 || target >= items.length) {
      return;
    }
    const [item] = items.splice(index, 1);
    items.splice(target, 0, item!);
    await persistList({
      ...activeList,
      clientId: selected.resource.id,
      scheduleId: activeRequestScheduleId,
      requestItems: items,
    });
  }

  function openEditRequestItem(itemId: string, label: string) {
    setEditingRequestItemId(itemId);
    setEditRequestLabel(label);
  }

  function cancelEditRequestItem() {
    setEditingRequestItemId(null);
    setEditRequestLabel("");
  }

  async function handleSaveEditRequestItem() {
    if (!selected || !activeRequestScheduleId || !editingRequestItemId) {
      return;
    }
    const label = editRequestLabel.trim();
    if (!label) {
      return;
    }
    await persistList(
      updateRequestBulletLabel(
        {
          ...activeList,
          clientId: selected.resource.id,
          scheduleId: activeRequestScheduleId,
        },
        editingRequestItemId,
        label,
      ),
    );
    cancelEditRequestItem();
  }

  async function handleRemoveRequestItem(itemId: string, label: string) {
    if (!selected || !activeRequestScheduleId) {
      return;
    }
    if (!window.confirm(`Delete "${label}"?`)) {
      return;
    }
    if (editingRequestItemId === itemId) {
      cancelEditRequestItem();
    }
    await persistList(
      removeRequestBullet(
        {
          ...activeList,
          clientId: selected.resource.id,
          scheduleId: activeRequestScheduleId,
        },
        itemId,
      ),
    );
  }

  async function handleCompleteActionItem(
    itemId: string,
    status: ActionResolveTarget = "completed",
  ) {
    if (!selected) {
      return;
    }
    const owner =
      findListContainingActionItem(
        requestLists,
        selected.resource.id,
        itemId,
      ) ?? activeList;
    await persistList(completeActionItem(owner, itemId, status));
  }

  async function handleReopenActionItem(itemId: string) {
    if (!selected) {
      return;
    }
    const owner =
      findListContainingActionItem(
        requestLists,
        selected.resource.id,
        itemId,
      ) ?? activeList;
    await persistList(reopenActionItem(owner, itemId));
  }

  async function resolveMailAccount() {
    let account =
      instance.getActiveAccount() ??
      instance.getAllAccounts()[0] ??
      msalAccount;

    if (!account) {
      const hint = emailFrom || email?.trim() || "";
      setEmailSendMessage("Connecting your Microsoft account to send…");
      const result = await instance.loginPopup({
        ...mailSendRequest,
        loginHint: hint || undefined,
        prompt: "select_account",
      });
      account = result.account;
      if (account) {
        instance.setActiveAccount(account);
      }
    }

    if (!account) {
      throw new Error(
        "Microsoft sign-in was cancelled. Sign in with the Microsoft account that owns this mailbox, then try again.",
      );
    }

    return account;
  }

  function formatMailSendError(sendError: unknown): string {
    const raw =
      sendError instanceof Error ? sendError.message : "Failed to send email.";
    const lower = raw.toLowerCase();
    if (
      lower.includes("aadsts50020") ||
      lower.includes("does not exist in tenant")
    ) {
      return (
        "sana@epsfresno.com is not in the Fire Fly Azure tenant. In Azure Portal, set this app to multi-tenant (any organizational directory), restart the app, then try again and sign in with sana@epsfresno.com."
      );
    }
    return raw;
  }

  async function handleSendEmailNow() {
    if (!emailTo.trim() || !emailSubject.trim()) {
      setEmailSendError("To and Subject are required.");
      setEmailSendMessage(null);
      return;
    }

    setSendingEmail(true);
    setEmailSendError(null);
    setEmailSendMessage(null);
    try {
      const account = await resolveMailAccount();
      setEmailSendMessage(null);
      await sendGraphMail(instance, account, {
        to: emailTo,
        cc: emailCc,
        bcc: emailBcc,
        subject: emailSubject,
        bodyText: emailBody,
        highlights: emailAutofillHighlights,
        portalHref: emailPortalHref,
        signature: emailSignature,
      });
      setEmailSendMessage("Email sent.");
    } catch (sendError) {
      setEmailSendError(formatMailSendError(sendError));
      setEmailSendMessage(null);
    } finally {
      setSendingEmail(false);
    }
  }

  async function handleScheduleEmail() {
    if (!emailTo.trim() || !emailSubject.trim()) {
      setEmailSendError("To and Subject are required.");
      setEmailSendMessage(null);
      return;
    }

    const sendAt = parseDateTimeLocalValue(emailSendAt);
    if (!sendAt) {
      setEmailSendError("Choose a valid Send at date and time.");
      setEmailSendMessage(null);
      return;
    }

    setSchedulingEmail(true);
    setEmailSendError(null);
    setEmailSendMessage(null);
    try {
      const account = await resolveMailAccount();
      setEmailSendMessage(null);
      await sendGraphMail(instance, account, {
        to: emailTo,
        cc: emailCc,
        bcc: emailBcc,
        subject: emailSubject,
        bodyText: emailBody,
        highlights: emailAutofillHighlights,
        portalHref: emailPortalHref,
        signature: emailSignature,
        sendAt,
      });
      setEmailSendMessage(
        `Email scheduled for ${formatBusinessDayLabel(sendAt)} at ${sendAt.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}.`,
      );
    } catch (sendError) {
      setEmailSendError(formatMailSendError(sendError));
      setEmailSendMessage(null);
    } finally {
      setSchedulingEmail(false);
    }
  }

  return (
    <section className="content-panel content-panel--actions">
      {error ? <p className="form-message error">{error}</p> : null}

      {!ready ? (
        <p className="table-empty">Loading…</p>
      ) : assignedClients.length === 0 ? (
        <p className="table-empty">
          No clients are assigned to you in Resource Allocation yet.
        </p>
      ) : (
        <div className="client-management-stack">
          <div className="client-picker">
            <div
              className="client-picker-list"
              role="tablist"
              aria-label="Assigned clients"
            >
              {assignedClients.map((entry) => {
                const isActive = selected?.resource.id === entry.resource.id;
                return (
                  <button
                    key={entry.resource.id}
                    type="button"
                    role="tab"
                    aria-selected={isActive}
                    className={
                      isActive
                        ? "client-picker-chip active"
                        : "client-picker-chip"
                    }
                    onClick={() => setSelectedId(entry.resource.id)}
                  >
                    {entry.resource.clientName || "Untitled client"}
                  </button>
                );
              })}
            </div>
          </div>

          <div className="client-management-panel">
            {selected ? (
              <>
                <div className="client-info-section">
                  <button
                    type="button"
                    className={
                      clientInfoOpen
                        ? "client-info-accordion-trigger open"
                        : "client-info-accordion-trigger"
                    }
                    aria-expanded={clientInfoOpen}
                    onClick={() => setClientInfoOpen((open) => !open)}
                  >
                    <span className="client-info-accordion-label">
                      Client Info
                    </span>
                    <span
                      className="client-info-accordion-chevron"
                      aria-hidden="true"
                    >
                      ▾
                    </span>
                  </button>
                  {clientInfoOpen ? (
                    <div className="table-wrap client-info-table-wrap">
                      <table className="projects-table client-info-table">
                        <tbody>
                          {buildClientInfoRows(selected).map(
                            ([left, right]) => (
                              <tr
                                key={`${left.key}-${right?.key ?? "empty"}`}
                              >
                                <th scope="row" className="client-info-field">
                                  {left.label}
                                </th>
                                <td className="client-info-value">
                                  {left.value}
                                </td>
                                <th scope="row" className="client-info-field">
                                  {right?.label ?? ""}
                                </th>
                                <td className="client-info-value">
                                  {right?.value ?? ""}
                                </td>
                              </tr>
                            ),
                          )}
                        </tbody>
                      </table>
                    </div>
                  ) : null}
                </div>

                <div className="adjustable-settings-section">
                  <button
                    type="button"
                    className={
                      adjustableSettingsOpen
                        ? "client-info-accordion-trigger open"
                        : "client-info-accordion-trigger"
                    }
                    aria-expanded={adjustableSettingsOpen}
                    onClick={() =>
                      setAdjustableSettingsOpen((open) => !open)
                    }
                  >
                    <span className="client-info-accordion-label">
                      Adjustable Setting
                    </span>
                    <span
                      className="client-info-accordion-chevron"
                      aria-hidden="true"
                    >
                      ▾
                    </span>
                  </button>
                  {adjustableSettingsOpen ? (
                    <div className="adjustable-settings-body">
                      <div className="adjustable-settings-header">
                        <p className="lookup-hint">
                          Map an email template to a business day for this
                          client. Check one or more rows, then use Delete /
                          Edit / Add.
                        </p>
                        {canEdit ? (
                          <div className="adjustable-settings-actions">
                            {adjustableDraft ? (
                              <>
                                <button
                                  type="button"
                                  className="button primary"
                                  disabled={savingAdjustable}
                                  onClick={() => void saveAdjustableDraft()}
                                >
                                  {savingAdjustable ? "Saving…" : "Save"}
                                </button>
                                <button
                                  type="button"
                                  className="button secondary"
                                  disabled={savingAdjustable}
                                  onClick={() => setAdjustableDraft(null)}
                                >
                                  Cancel
                                </button>
                              </>
                            ) : (
                              <>
                                <button
                                  type="button"
                                  className="button danger"
                                  disabled={
                                    savingAdjustable ||
                                    selectedAdjustableRowIds.length === 0
                                  }
                                  onClick={() =>
                                    void deleteSelectedAdjustableRows()
                                  }
                                >
                                  Delete
                                </button>
                                <button
                                  type="button"
                                  className="button secondary-solid"
                                  disabled={
                                    savingAdjustable || !selectedAdjustableRow
                                  }
                                  onClick={startEditSelectedAdjustableRow}
                                >
                                  Edit
                                </button>
                                <button
                                  type="button"
                                  className="button primary"
                                  disabled={savingAdjustable}
                                  onClick={startAddAdjustableRow}
                                >
                                  Add
                                </button>
                              </>
                            )}
                          </div>
                        ) : null}
                      </div>
                      <div className="table-wrap">
                        <table className="projects-table projects-table--with-check adjustable-settings-table">
                          <thead>
                            <tr>
                              <th className="col-check" scope="col">
                                <input
                                  type="checkbox"
                                  checked={allAdjustableSelected}
                                  disabled={
                                    adjustableRows.length === 0 ||
                                    Boolean(adjustableDraft)
                                  }
                                  onChange={toggleSelectAllAdjustableRows}
                                  aria-label="Select all adjustable settings"
                                />
                              </th>
                              <th scope="col">Template</th>
                              <th scope="col">Business Day</th>
                            </tr>
                          </thead>
                          <tbody>
                            {adjustableRows.map((row) => {
                              const isEditing = adjustableDraft?.id === row.id;
                              const isSelected =
                                selectedAdjustableRowIds.includes(row.id);
                              return (
                                <tr
                                  key={row.id}
                                  className={
                                    isSelected
                                      ? "row-selected row-selectable"
                                      : "row-selectable"
                                  }
                                  onClick={() => {
                                    if (adjustableDraft) {
                                      return;
                                    }
                                    toggleAdjustableRowSelected(row.id);
                                  }}
                                >
                                  <td className="col-check">
                                    <input
                                      type="checkbox"
                                      checked={isSelected}
                                      disabled={Boolean(adjustableDraft)}
                                      onChange={() =>
                                        toggleAdjustableRowSelected(row.id)
                                      }
                                      onClick={(event) =>
                                        event.stopPropagation()
                                      }
                                      aria-label={`Select ${templateLabelForId(row.templateId)}`}
                                    />
                                  </td>
                                  <td>
                                    {isEditing && adjustableDraft ? (
                                      <select
                                        className="field-input"
                                        value={adjustableDraft.templateId}
                                        onClick={(event) =>
                                          event.stopPropagation()
                                        }
                                        onChange={(event) =>
                                          setAdjustableDraft({
                                            ...adjustableDraft,
                                            templateId: event.target.value,
                                          })
                                        }
                                      >
                                        {emailTemplates.templates.map(
                                          (template) => (
                                            <option
                                              key={template.id}
                                              value={template.id}
                                            >
                                              {templateOptionLabel(template)}
                                            </option>
                                          ),
                                        )}
                                      </select>
                                    ) : (
                                      templateLabelForId(row.templateId)
                                    )}
                                  </td>
                                  <td>
                                    {isEditing && adjustableDraft ? (
                                      <select
                                        className="field-input"
                                        value={adjustableDraft.businessDayId}
                                        onClick={(event) =>
                                          event.stopPropagation()
                                        }
                                        onChange={(event) =>
                                          setAdjustableDraft({
                                            ...adjustableDraft,
                                            businessDayId: event.target.value,
                                          })
                                        }
                                      >
                                        {CLIENT_SCHEDULE_ITEMS.map((item) => (
                                          <option
                                            key={item.id}
                                            value={item.id}
                                          >
                                            {item.label}
                                          </option>
                                        ))}
                                      </select>
                                    ) : (
                                      businessDayLabel(row.businessDayId)
                                    )}
                                  </td>
                                </tr>
                              );
                            })}
                            {adjustableDraft && adjustableDraft.id === null ? (
                              <tr className="row-selected">
                                <td className="col-check" />
                                <td>
                                  <select
                                    className="field-input"
                                    value={adjustableDraft.templateId}
                                    onChange={(event) =>
                                      setAdjustableDraft({
                                        ...adjustableDraft,
                                        templateId: event.target.value,
                                      })
                                    }
                                  >
                                    {emailTemplates.templates.length === 0 ? (
                                      <option value="">No templates</option>
                                    ) : null}
                                    {emailTemplates.templates.map(
                                      (template) => (
                                        <option
                                          key={template.id}
                                          value={template.id}
                                        >
                                          {templateOptionLabel(template)}
                                        </option>
                                      ),
                                    )}
                                  </select>
                                </td>
                                <td>
                                  <select
                                    className="field-input"
                                    value={adjustableDraft.businessDayId}
                                    onChange={(event) =>
                                      setAdjustableDraft({
                                        ...adjustableDraft,
                                        businessDayId: event.target.value,
                                      })
                                    }
                                  >
                                    {CLIENT_SCHEDULE_ITEMS.map((item) => (
                                      <option key={item.id} value={item.id}>
                                        {item.label}
                                      </option>
                                    ))}
                                  </select>
                                </td>
                              </tr>
                            ) : null}
                            {adjustableRows.length === 0 &&
                            !adjustableDraft ? (
                              <tr>
                                <td colSpan={3} className="table-empty">
                                  No adjustable settings yet.
                                  {canEdit
                                    ? " Click Add to map a template to a business day."
                                    : ""}
                                </td>
                              </tr>
                            ) : null}
                          </tbody>
                        </table>
                      </div>
                      {adjustableMessage ? (
                        <p className="form-message success">
                          {adjustableMessage}
                        </p>
                      ) : null}
                      {adjustableError ? (
                        <p className="form-message error">{adjustableError}</p>
                      ) : null}
                    </div>
                  ) : null}
                </div>

                <div className="client-schedule-layout">
                  <aside
                    className="lookup-fields client-schedule-sidebar"
                    aria-label="Client schedule"
                  >
                    {sidebarScheduleItems.length === 0 ? (
                      <p className="lookup-hint">
                        Add a Template + Business Day above to show schedules
                        here.
                      </p>
                    ) : (
                      sidebarScheduleItems.map((item) => {
                        const businessDate = scheduleDateById.get(item.id);
                        const nativeCount = sidebarTemplateItems.filter(
                          (entry) => entry.businessDayId === item.id,
                        ).length;
                        const followUpCount = templateFollowUps.filter(
                          (entry) => entry.followUpBusinessDayId === item.id,
                        ).length;
                        const templateCount = nativeCount + followUpCount;
                        return (
                          <button
                            key={item.id}
                            type="button"
                            className={
                              selectedScheduleId === item.id
                                ? "lookup-field-link active"
                                : "lookup-field-link"
                            }
                            onClick={() => {
                              setSelectedScheduleId(item.id);
                              const firstOnBd = sidebarTemplateItems.find(
                                (entry) => entry.businessDayId === item.id,
                              );
                              if (firstOnBd) {
                                setSelectedSidebarRowId(firstOnBd.rowId);
                                setRequestTypeId(
                                  requestTypeForTemplateId(
                                    firstOnBd.templateId,
                                  ),
                                );
                                setWorkflowTab(
                                  firstOnBd.templateId === "template-1" ||
                                    firstOnBd.templateId === "template-2"
                                    ? "request"
                                    : "action",
                                );
                                return;
                              }
                              const followUp = templateFollowUps.find(
                                (entry) =>
                                  entry.followUpBusinessDayId === item.id,
                              );
                              if (followUp) {
                                setSelectedSidebarRowId(followUp.sourceRowId);
                                setRequestTypeId(
                                  requestTypeForTemplateId(followUp.templateId),
                                );
                                setWorkflowTab("action");
                              }
                            }}
                          >
                            <span className="client-schedule-link-text">
                              <span className="client-schedule-link-label">
                                {item.label}
                              </span>
                              {businessDate ? (
                                <span className="client-schedule-link-date">
                                  {businessDate}
                                  {templateCount > 1
                                    ? ` · ${templateCount} templates`
                                    : ""}
                                </span>
                              ) : null}
                            </span>
                          </button>
                        );
                      })
                    )}
                  </aside>

                  <div className="client-schedule-panel">
                    {sidebarScheduleItems.length === 0 ? (
                      <p className="table-empty">
                        No business days configured for this client yet. Open
                        Adjustable Setting and add at least one row.
                      </p>
                    ) : (
                      <>
                    <div className="client-schedule-panel-header">
                      <h3 className="section-title">
                        {selectedSchedule?.label}
                        {selectedScheduleDate
                          ? ` — ${selectedScheduleDate}`
                          : ""}
                      </h3>
                    </div>

                    {(() => {
                      const workflowTabsClassName = showTemplateChooser
                        ? "lookup-section-tabs bd1-workflow-tabs request-workflow-tabs"
                        : "lookup-section-tabs request-workflow-tabs";

                      const requestWorkflow = (
                        <>
                          {showBd1LikeWorkflow ? (
                            <p className="bd1-workflow-kicker">
                              {showBd7Content
                                ? isTemplateLateFollowUpDay
                                  ? `Late Action Items shows only items marked Late on the +4 review day for ${isTemplate2 ? "Template 2" : "Template 1"}.`
                                  : "Late Action Items shows only items marked Late on BD5 for the selected request type."
                                : "Use the tabs below for request list, action items, and email."}
                            </p>
                          ) : null}
                          <div
                            className={workflowTabsClassName}
                            role="tablist"
                            aria-label="Request workflow"
                          >
                      {showRequestListTab ? (
                        <button
                          type="button"
                          role="tab"
                          aria-selected={workflowTab === "request"}
                          className={
                            workflowTab === "request"
                              ? "lookup-section-tab active"
                              : "lookup-section-tab"
                          }
                          onClick={() => setWorkflowTab("request")}
                        >
                          Request List Template
                        </button>
                      ) : null}
                      <button
                        type="button"
                        role="tab"
                        aria-selected={workflowTab === "action"}
                        className={
                          workflowTab === "action"
                            ? "lookup-section-tab active"
                            : "lookup-section-tab"
                        }
                        onClick={() => setWorkflowTab("action")}
                      >
                        {showBd7Content ? "Late Action Items" : "Action Items"}
                      </button>
                      {showBd1LikeWorkflow ? (
                        <button
                          type="button"
                          role="tab"
                          aria-selected={workflowTab === "email"}
                          className={
                            workflowTab === "email"
                              ? "lookup-section-tab active"
                              : "lookup-section-tab"
                          }
                          onClick={() => setWorkflowTab("email")}
                        >
                          Email
                        </button>
                      ) : null}
                      {!showBd1LikeWorkflow && !showBd5Content ? (
                        <button
                          type="button"
                          role="tab"
                          aria-selected={workflowTab === "completed"}
                          className={
                            workflowTab === "completed"
                              ? "lookup-section-tab active"
                              : "lookup-section-tab"
                          }
                          onClick={() => setWorkflowTab("completed")}
                        >
                          Completed
                        </button>
                      ) : null}
                    </div>

                    {showRequestListTab && workflowTab === "request" ? (
                      <div className="request-list-panel">
                        <p className="lookup-hint">
                          Add items here. Items created this month appear on
                          Action Items right away; the full list also rolls
                          into Action Items when a new month starts.
                        </p>
                        <div className="request-list-add">
                          <input
                            className="field-input"
                            value={draftRequestLabel}
                            placeholder="Add bullet point…"
                            disabled={savingRequestList}
                            onChange={(event) =>
                              setDraftRequestLabel(event.target.value)
                            }
                            onKeyDown={(event) => {
                              if (event.key === "Enter") {
                                event.preventDefault();
                                void handleAddRequestItem();
                              }
                            }}
                          />
                          <button
                            type="button"
                            className="button secondary-solid"
                            disabled={
                              savingRequestList || !draftRequestLabel.trim()
                            }
                            onClick={() => void handleAddRequestItem()}
                          >
                            Add
                          </button>
                        </div>

                        {activeList.requestItems.length === 0 ? (
                          <p className="lookup-hint">
                            No request bullets yet.
                          </p>
                        ) : (
                          <ul className="request-list-items">
                            {activeList.requestItems.map((item, index) => {
                              const isEditing =
                                editingRequestItemId === item.id;
                              return (
                                <li
                                  key={item.id}
                                  className="request-list-item"
                                >
                                  {isEditing ? (
                                    <div className="request-list-edit">
                                      <input
                                        className="field-input"
                                        value={editRequestLabel}
                                        disabled={savingRequestList}
                                        autoFocus
                                        onChange={(event) =>
                                          setEditRequestLabel(
                                            event.target.value,
                                          )
                                        }
                                        onKeyDown={(event) => {
                                          if (event.key === "Enter") {
                                            event.preventDefault();
                                            void handleSaveEditRequestItem();
                                          }
                                          if (event.key === "Escape") {
                                            event.preventDefault();
                                            cancelEditRequestItem();
                                          }
                                        }}
                                      />
                                      <button
                                        type="button"
                                        className="button secondary-solid"
                                        disabled={
                                          savingRequestList ||
                                          !editRequestLabel.trim()
                                        }
                                        onClick={() =>
                                          void handleSaveEditRequestItem()
                                        }
                                      >
                                        Save
                                      </button>
                                      <button
                                        type="button"
                                        className="button secondary"
                                        disabled={savingRequestList}
                                        onClick={cancelEditRequestItem}
                                      >
                                        Cancel
                                      </button>
                                    </div>
                                  ) : (
                                    <>
                                      <span className="request-list-bullet">
                                        {item.label}
                                      </span>
                                      <div
                                        className="holiday-actions"
                                        aria-label="Action"
                                      >
                                        <button
                                          type="button"
                                          className="holiday-icon-btn"
                                          aria-label="Move up"
                                          title="Move up"
                                          disabled={
                                            index === 0 || savingRequestList
                                          }
                                          onClick={() =>
                                            void handleMoveRequestItem(
                                              item.id,
                                              -1,
                                            )
                                          }
                                        >
                                          ↑
                                        </button>
                                        <button
                                          type="button"
                                          className="holiday-icon-btn"
                                          aria-label="Move down"
                                          title="Move down"
                                          disabled={
                                            index ===
                                              activeList.requestItems.length -
                                                1 || savingRequestList
                                          }
                                          onClick={() =>
                                            void handleMoveRequestItem(
                                              item.id,
                                              1,
                                            )
                                          }
                                        >
                                          ↓
                                        </button>
                                        <button
                                          type="button"
                                          className="holiday-icon-btn"
                                          aria-label="Edit"
                                          title="Edit"
                                          disabled={savingRequestList}
                                          onClick={() =>
                                            openEditRequestItem(
                                              item.id,
                                              item.label,
                                            )
                                          }
                                        >
                                          ✎
                                        </button>
                                        <button
                                          type="button"
                                          className="holiday-icon-btn holiday-icon-btn--danger"
                                          aria-label="Delete"
                                          title="Delete"
                                          disabled={savingRequestList}
                                          onClick={() =>
                                            void handleRemoveRequestItem(
                                              item.id,
                                              item.label,
                                            )
                                          }
                                        >
                                          ×
                                        </button>
                                      </div>
                                    </>
                                  )}
                                </li>
                              );
                            })}
                          </ul>
                        )}
                      </div>
                    ) : null}

                    {workflowTab === "action" ? (
                      <div className="request-list-panel">
                        {showBd5Content ? (
                          <>
                            <p className="lookup-hint">
                              {isTemplateReviewFollowUpDay
                                ? `${isTemplate2 ? "Template 2" : "Template 1"} action items (send business day + 4). Mark each row Completed or Late.`
                                : "BD1 action items for this request type. Mark each row Completed or Late."}
                            </p>
                            {bd5ActionTableItems.length === 0 ? (
                              <p className="lookup-hint">
                                No action items for this type yet.
                              </p>
                            ) : (
                              <div className="table-wrap bd5-action-table-wrap">
                                <table className="projects-table bd5-action-table">
                                  <thead>
                                    <tr>
                                      <th scope="col">List</th>
                                      <th scope="col">Date Time</th>
                                      <th scope="col">Status</th>
                                      <th scope="col">Action</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {bd5ActionTableItems.map((item) => {
                                      const resolvedAt =
                                        formatActionResolvedAt(
                                          item.completedAt,
                                        );
                                      const isOpen = item.status === "action";
                                      const statusLabel = isOpen
                                        ? "Open"
                                        : item.status === "late"
                                          ? "Late"
                                          : "Completed";
                                      const statusClass = isOpen
                                        ? "bd5-action-status bd5-action-status--open"
                                        : item.status === "late"
                                          ? "bd5-action-status bd5-action-status--late"
                                          : "bd5-action-status bd5-action-status--completed";
                                      return (
                                        <tr
                                          key={item.id}
                                          className={
                                            item.status === "late"
                                              ? "bd5-action-row bd5-action-row--late"
                                              : item.status === "completed"
                                                ? "bd5-action-row bd5-action-row--completed"
                                                : "bd5-action-row"
                                          }
                                        >
                                          <td className="bd5-action-list-cell">
                                            {item.label}
                                          </td>
                                          <td className="bd5-action-datetime-cell">
                                            {resolvedAt || "—"}
                                          </td>
                                          <td>
                                            <span className={statusClass}>
                                              {statusLabel}
                                            </span>
                                          </td>
                                          <td>
                                            {isOpen ? (
                                              <div className="bd5-action-row-actions">
                                                <button
                                                  type="button"
                                                  className="button secondary-solid bd5-action-btn bd5-action-btn--completed"
                                                  disabled={savingRequestList}
                                                  onClick={() =>
                                                    void handleCompleteActionItem(
                                                      item.id,
                                                      "completed",
                                                    )
                                                  }
                                                >
                                                  Completed
                                                </button>
                                                <button
                                                  type="button"
                                                  className="button secondary-solid bd5-action-btn bd5-action-btn--late"
                                                  disabled={savingRequestList}
                                                  onClick={() =>
                                                    void handleCompleteActionItem(
                                                      item.id,
                                                      "late",
                                                    )
                                                  }
                                                >
                                                  Late
                                                </button>
                                              </div>
                                            ) : (
                                              <button
                                                type="button"
                                                className="button secondary"
                                                disabled={savingRequestList}
                                                onClick={() =>
                                                  void handleReopenActionItem(
                                                    item.id,
                                                  )
                                                }
                                              >
                                                Reopen
                                              </button>
                                            )}
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            )}
                          </>
                        ) : showBd7Content ? (
                          <>
                            <p className="lookup-hint">
                              {isTemplateLateFollowUpDay
                                ? `Items marked Late on the +4 review day for ${isTemplate2 ? "Template 2" : "Template 1"} (send business day + 6).`
                                : "Items marked Late on BD5 for this category (Bank and Credit Card, or Other Statements)."}
                            </p>
                            {bd7LateItems.length === 0 ? (
                              <p className="lookup-hint">
                                No late action items for this category yet.
                              </p>
                            ) : (
                              <div className="table-wrap bd5-action-table-wrap">
                                <table className="projects-table bd5-action-table">
                                  <thead>
                                    <tr>
                                      <th scope="col">List</th>
                                      <th scope="col">Date Time</th>
                                      <th scope="col">Status</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {bd7LateItems.map((item) => {
                                      const resolvedAt =
                                        formatActionResolvedAt(
                                          item.completedAt,
                                        );
                                      return (
                                        <tr
                                          key={item.id}
                                          className="bd5-action-row bd5-action-row--late"
                                        >
                                          <td className="bd5-action-list-cell">
                                            {item.label}
                                          </td>
                                          <td className="bd5-action-datetime-cell">
                                            {resolvedAt || "—"}
                                          </td>
                                          <td>
                                            <span className="bd5-action-status bd5-action-status--late">
                                              Late
                                            </span>
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            )}
                          </>
                        ) : (
                          <>
                            <p className="lookup-hint">
                              Includes Request List items created this month,
                              plus items carried over at month start. Check an
                              item to mark it completed or late.
                            </p>
                            {actionItems.length === 0 ? (
                              <p className="lookup-hint">
                                No action items for this month yet.
                              </p>
                            ) : (
                              <ul className="request-list-items">
                                {actionItems.map((item) => (
                                  <li
                                    key={item.id}
                                    className="request-list-item"
                                  >
                                    <label className="request-list-check">
                                      <input
                                        type="checkbox"
                                        checked={false}
                                        disabled={savingRequestList}
                                        onChange={() =>
                                          void handleCompleteActionItem(
                                            item.id,
                                            "completed",
                                          )
                                        }
                                      />
                                      <span className="request-list-label">
                                        {item.label}
                                      </span>
                                    </label>
                                  </li>
                                ))}
                              </ul>
                            )}
                          </>
                        )}
                      </div>
                    ) : null}

                    {!showBd1LikeWorkflow &&
                    !showBd5Content &&
                    workflowTab === "completed" ? (
                      <div className="request-list-panel">
                        <p className="lookup-hint">
                          Uncheck an item to move it back to Action Items.
                        </p>
                        {completedItems.length === 0 ? (
                          <p className="lookup-hint">
                            No completed items yet.
                          </p>
                        ) : (
                          <ul className="request-list-items">
                            {completedItems.map((item) => {
                              const resolvedAt = formatActionResolvedAt(
                                item.completedAt,
                              );
                              return (
                                <li key={item.id} className="request-list-item">
                                  <label className="request-list-check">
                                    <input
                                      type="checkbox"
                                      checked
                                      disabled={savingRequestList}
                                      onChange={() =>
                                        void handleReopenActionItem(item.id)
                                      }
                                    />
                                    <span className="request-list-label-block">
                                      <span className="request-list-label checked">
                                        {item.label}
                                      </span>
                                      {resolvedAt ? (
                                        <span className="request-list-resolved-at">
                                          {resolvedAt}
                                        </span>
                                      ) : null}
                                    </span>
                                  </label>
                                </li>
                              );
                            })}
                          </ul>
                        )}
                      </div>
                    ) : null}

                    {showBd1LikeWorkflow && workflowTab === "email" ? (
                      <div className="client-schedule-email client-schedule-email--tab">
                        <h4 className="client-schedule-email-title">Email</h4>
                        <p className="lookup-hint">
                          Subject and body are filled from{" "}
                          {selectedSidebarItem?.label ?? "the selected template"}
                          .
                        </p>
                        <div className="client-schedule-email-fields">
                          <label className="field">
                            <span className="field-label">To</span>
                            <input
                              className="field-input"
                              type="text"
                              inputMode="email"
                              value={emailTo}
                              onChange={(event) =>
                                setEmailTo(event.target.value)
                              }
                              placeholder="email@example.com, email2@example.com"
                            />
                          </label>
                          <label className="field">
                            <span className="field-label">From</span>
                            <input
                              className="field-input field-input--readonly"
                              type="email"
                              value={emailFrom}
                              readOnly
                              tabIndex={-1}
                              aria-readonly="true"
                              title="From is the Microsoft account you signed in with"
                            />
                          </label>
                          <label className="field">
                            <span className="field-label">CC</span>
                            <input
                              className="field-input"
                              type="text"
                              inputMode="email"
                              value={emailCc}
                              onChange={(event) =>
                                setEmailCc(event.target.value)
                              }
                              placeholder="email@example.com, email2@example.com"
                            />
                          </label>
                          <label className="field">
                            <span className="field-label">BCC</span>
                            <input
                              className="field-input"
                              type="text"
                              inputMode="email"
                              value={emailBcc}
                              onChange={(event) =>
                                setEmailBcc(event.target.value)
                              }
                              placeholder="email@example.com, email2@example.com"
                            />
                          </label>
                          <label className="field">
                            <span className="field-label">Deadline</span>
                            <input
                              className="field-input"
                              type="date"
                              value={effectiveEmailDeadline}
                              onChange={(event) =>
                                setEmailDeadline(event.target.value)
                              }
                            />
                          </label>
                          <label className="field">
                            <span className="field-label">Send at</span>
                            <input
                              className="field-input"
                              type="datetime-local"
                              value={emailSendAt}
                              onChange={(event) =>
                                setEmailSendAt(event.target.value)
                              }
                            />
                          </label>
                          <div className="field field--wide">
                            <span className="field-label">Subject</span>
                            <div className="field-input field-input--readonly client-schedule-email-rich">
                              {renderAutofilledText(
                                emailSubject,
                                emailAutofillHighlights,
                                emailPortalHref,
                              ) || "—"}
                            </div>
                          </div>
                          <div className="field field--wide">
                            <span className="field-label">Body</span>
                            <div className="field-input field-input--readonly email-template-body-input client-schedule-email-rich client-schedule-email-rich--body">
                              {emailBody.trim() ? (
                                <>
                                  {renderAutofilledText(
                                    emailBody,
                                    emailAutofillHighlights,
                                    emailPortalHref,
                                  )}
                                  {"\n\n"}
                                  <span>Thank you,</span>
                                  {"\n\n"}
                                  {hasEmailSignature(emailSignature) ? (
                                    <ClientScheduleEmailSignature
                                      signature={emailSignature}
                                    />
                                  ) : (
                                    <p className="lookup-hint email-signature-hint">
                                      Set your email signature under My Profile
                                      to show it after Thank you.
                                    </p>
                                  )}
                                </>
                              ) : (
                                "—"
                              )}
                            </div>
                          </div>
                        </div>
                        <div className="client-schedule-email-actions">
                          <button
                            type="button"
                            className="button primary"
                            disabled={
                              sendingEmail ||
                              schedulingEmail ||
                              !emailTo.trim() ||
                              !emailSubject.trim()
                            }
                            onClick={() => void handleSendEmailNow()}
                          >
                            {sendingEmail ? "Sending…" : "Send Now"}
                          </button>
                          <button
                            type="button"
                            className="button secondary-solid"
                            disabled={
                              sendingEmail ||
                              schedulingEmail ||
                              !emailTo.trim() ||
                              !emailSubject.trim() ||
                              !emailSendAt.trim()
                            }
                            onClick={() => void handleScheduleEmail()}
                          >
                            {schedulingEmail ? "Scheduling…" : "Schedule"}
                          </button>
                        </div>
                        {emailSendMessage ? (
                          <p className="form-message success">
                            {emailSendMessage}
                          </p>
                        ) : null}
                        {emailSendError ? (
                          <p className="form-message error">{emailSendError}</p>
                        ) : null}
                      </div>
                    ) : null}
                        </>
                      );

                      if (!showTemplateChooser && !isTemplateFollowUpDay) {
                        return requestWorkflow;
                      }

                      return (
                        <div className="bd1-request-type-block">
                          {showTemplateChooser ? (
                            <>
                          <p className="bd1-request-type-kicker">
                            Step 1 — Choose template
                          </p>
                          <div
                            className="bd1-request-choices"
                            role="tablist"
                            aria-label={`${selectedSchedule?.label ?? "Schedule"} templates`}
                          >
                            {templatesForSelectedBd.map((item) => {
                              const isActive =
                                selectedSidebarItem?.rowId === item.rowId;
                              const isFollowUpChoice = templateFollowUps.some(
                                (entry) =>
                                  entry.sourceRowId === item.rowId &&
                                  entry.followUpBusinessDayId ===
                                    selectedScheduleId,
                              );
                              return (
                                <button
                                  key={item.rowId}
                                  type="button"
                                  role="tab"
                                  aria-selected={isActive}
                                  className={
                                    isActive
                                      ? "bd1-request-choice bd1-request-choice--active"
                                      : "bd1-request-choice"
                                  }
                                  onClick={() => {
                                    setSelectedSidebarRowId(item.rowId);
                                    setRequestTypeId(
                                      requestTypeForTemplateId(item.templateId),
                                    );
                                    setWorkflowTab(
                                      isFollowUpChoice
                                        ? "action"
                                        : item.templateId === "template-1" ||
                                            item.templateId === "template-2"
                                          ? "request"
                                          : "action",
                                    );
                                  }}
                                >
                                  <span className="bd1-request-choice-title">
                                    {item.label}
                                  </span>
                                  <span
                                    className={
                                      isActive
                                        ? "bd1-request-choice-status"
                                        : "bd1-request-choice-status bd1-request-choice-status--muted"
                                    }
                                  >
                                    {isActive ? "Selected" : "Select"}
                                  </span>
                                </button>
                              );
                            })}
                          </div>
                            </>
                          ) : null}

                          <div
                            className="bd1-request-body"
                            role="tabpanel"
                            aria-labelledby="bd1-request-body-title"
                          >
                            <div className="bd1-request-body-header">
                              <div className="bd1-request-body-header-main">
                                <span className="bd1-request-body-eyebrow">
                                  Step 2 — Content for this type
                                </span>
                                <h4
                                  id="bd1-request-body-title"
                                  className="bd1-request-body-title"
                                >
                                  {selectedSidebarItem?.label ??
                                    activeRequestType.label}
                                </h4>
                                <p className="bd1-request-body-hint">
                                  {isTemplateReviewFollowUpDay
                                    ? `Completed / Late review for ${isTemplate2 ? "Template 2" : "Template 1"} (send business day + 4).`
                                    : isTemplateLateFollowUpDay
                                      ? `Late Action Items for ${isTemplate2 ? "Template 2" : "Template 1"} (send business day + 6).`
                                      : "Everything below applies to this template only."}
                                </p>
                              </div>
                              <span
                                className="bd1-request-body-selected-pill"
                                aria-hidden="true"
                              >
                                Active
                              </span>
                            </div>
                            <div className="bd1-request-body-content">
                              <div className="bd1-workflow-section">
                                {requestWorkflow}
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })()}
                      </>
                    )}
                  </div>
                </div>
              </>
            ) : null}
          </div>
        </div>
      )}
    </section>
  );
}
