"use client";

import { useMsal } from "@azure/msal-react";
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type MouseEvent,
  type ReactNode,
} from "react";
import { mailSendRequest, sharePointRequest } from "@/auth/authConfig";
import { useAccess } from "@/auth/accessContext";
import { getAccountEmail } from "@/auth/organization";
import {
  formatBusinessDayLabel,
  getBd7DeadlineDate,
  getScheduleBusinessDate,
} from "@/lib/business-days";
import {
  CLIENT_MANAGEMENT_DEMO_MONTH_EVENT,
  getClientManagementNow,
  ROLL_CLIENT_MANAGEMENT_MONTH_ID,
  type RepeatedTasksStore,
} from "@/lib/repeated-tasks";
import {
  ASSIGNMENT_ROLE_LABELS,
  getClientsAssignedToUser,
  staffNameAliases,
  type AssignedClient,
} from "@/lib/client-management";
import {
  BD1_TEMPLATE_OPTIONS,
  bd1TemplateToRequestTypeId,
  normalizeBd1Template,
  normalizeClientProfile,
  type Bd1TemplateId,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import {
  addRequestBulletToList,
  completeActionItem,
  createEmptyRequestListsStore,
  currentMonthKey,
  findListContainingActionItem,
  getClientActionItems,
  getClientRequestList,
  normalizeClientRequestListsStore,
  removeRequestBullet,
  reopenActionItem,
  requestItemsForMonth,
  setActionItemStatus,
  updateRequestBulletLabel,
  upsertClientRequestList,
  type ActionWorkflowItem,
  type ClientRequestList,
  type ClientRequestListsStore,
} from "@/lib/client-request-lists";
import {
  listSharePointMonthFolderFiles,
  sharePointFileNameMatches,
} from "@/lib/sharepoint-graph";
import {
  createEmptyScheduleSettingsStore,
  getActiveScheduleIds,
  isScheduleActiveForClient,
  normalizeClientScheduleSettingsStore,
  toggleClientScheduleActive,
  type ClientScheduleSettingsStore,
} from "@/lib/client-schedule-settings";
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
  createDefaultHolidays,
  normalizeHolidays,
  type HolidaysStore,
} from "@/lib/holidays";
import { portalFetch } from "@/lib/portal-api";
import type { PortalUser } from "@/lib/portal-users";
import type { ResourceClientEntry } from "@/lib/resource-clients";
import {
  sendGraphMail,
  type GraphMailAttachment,
} from "@/lib/send-graph-mail";
import {
  CLIENT_TIME_ZONE_REQUIRED_MESSAGE,
  getClientNineAmSchedule,
  parseDateTimeLocalInTimeZone,
  resolveClientTimeZone,
  USER_TIME_ZONE,
} from "@/lib/time-zones";

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(
      ...bytes.subarray(offset, offset + chunkSize),
    );
  }
  return btoa(binary);
}

function mimeTypeForExcelFile(file: File): string {
  if (file.type.trim()) {
    return file.type.trim();
  }
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".csv")) {
    return "text/csv";
  }
  if (lower.endsWith(".xls")) {
    return "application/vnd.ms-excel";
  }
  return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
}

function mimeTypeForPdfFile(file: File): string {
  if (file.type.trim()) {
    return file.type.trim();
  }
  return "application/pdf";
}

function mimeTypeForBd15File(file: File): string {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".pdf") || file.type === "application/pdf") {
    return mimeTypeForPdfFile(file);
  }
  return mimeTypeForExcelFile(file);
}

type Bd15StoredAttachment = GraphMailAttachment & {
  id: string;
  savedAt: string;
};

type Bd15ApiAttachment = {
  id?: string;
  name?: string;
  contentType?: string;
  contentBytes?: string;
  savedAt?: string;
};

function mapBd15ApiAttachments(
  attachments: Bd15ApiAttachment[] | undefined,
): Bd15StoredAttachment[] {
  if (!Array.isArray(attachments)) {
    return [];
  }
  return attachments
    .filter(
      (item) =>
        typeof item.id === "string" &&
        item.id &&
        typeof item.name === "string" &&
        item.name &&
        typeof item.contentBytes === "string" &&
        item.contentBytes,
    )
    .map((item) => ({
      id: item.id as string,
      name: item.name as string,
      contentType: item.contentType || "application/pdf",
      contentBytes: item.contentBytes as string,
      savedAt:
        typeof item.savedAt === "string" && item.savedAt.trim()
          ? item.savedAt.trim()
          : "",
    }));
}

/**
 * Required BD15/BD18 attachments:
 * - Profit Loss: file name contains both "profit" and "loss" (any order/combo).
 * - Balance Sheet: file name contains the phrase "balance sheet" (one keyword).
 * Matching is always case-insensitive.
 */
const BD15_REQUIRED_REPORTS = [
  "Profit & Loss Account",
  "Balance Sheet",
] as const;

function normalizeBd15ReportKey(value: string): string {
  return value
    .toLowerCase()
    .replace(/\.[^.]+$/, "")
    .replace(/[/_\\-]+/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function bd15FileMatchesReport(
  fileName: string,
  requiredLabel: string,
): boolean {
  const fileKey = normalizeBd15ReportKey(fileName);
  if (!fileKey) {
    return false;
  }
  const reqKey = normalizeBd15ReportKey(requiredLabel);
  // Balance Sheet = one phrase keyword (case-insensitive).
  if (reqKey === "balance sheet") {
    return fileKey.includes("balance sheet");
  }
  // Profit & Loss Account = file name contains both "profit" and "loss".
  if (reqKey.includes("profit") && reqKey.includes("loss")) {
    return fileKey.includes("profit") && fileKey.includes("loss");
  }
  const requiredWords = reqKey.split(" ").filter(Boolean);
  if (requiredWords.length === 0) {
    return false;
  }
  return requiredWords.every((word) => fileKey.includes(word));
}

function getMissingBd15Reports(
  attachments: ReadonlyArray<{ name: string }>,
): string[] {
  return BD15_REQUIRED_REPORTS.filter(
    (required) =>
      !attachments.some((file) => bd15FileMatchesReport(file.name, required)),
  );
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
  { key: "bd1Template", label: "BD1 Template" },
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

/** Sidebar order: BD1→…→BD18, then end-of-month by date (Last 3 → Last 2 → Last 1), Client above Team. */
const CLIENT_SCHEDULE_ITEMS = [
  { id: "bd1", label: "BD1" },
  { id: "bd5", label: "BD5" },
  { id: "bd7", label: "BD7" },
  { id: "bd8", label: "BD8" },
  { id: "bd9", label: "BD9" },
  { id: "bd10", label: "BD10" },
  { id: "bd15", label: "BD15" },
  { id: "bd18", label: "BD18" },
  { id: "last-3-bd", label: "Last 3 BD" },
  { id: "last-2-bd", label: "Last 2 BD" },
  { id: "last-1-bd", label: "Last 1 BD" },
] as const;

const CLIENT_SCHEDULE_ORDER = new Map<string, number>(
  CLIENT_SCHEDULE_ITEMS.map((item, index) => [item.id, index]),
);

type LastBdAudience = "client" | "team";

function isLastBdSchedule(scheduleId: string): boolean {
  return /^last-\d+-bd$/i.test(scheduleId.trim());
}

function lastBdOrdinal(scheduleId: string): number | null {
  const match = /^last-(\d+)-bd$/i.exec(scheduleId.trim());
  if (!match) {
    return null;
  }
  const n = Number(match[1]);
  return Number.isFinite(n) && n > 0 ? n : null;
}

function lastBdLabel(scheduleId: string): string {
  const n = lastBdOrdinal(scheduleId);
  if (!n) {
    return "[Last BD Label]";
  }
  return n === 1
    ? "Last 1 business day of the month"
    : `Last ${n} business days of the month`;
}

const ALL_SCHEDULE_IDS = CLIENT_SCHEDULE_ITEMS.map((item) => item.id);

const BD1_SUB_MENUS = [
  {
    id: "bank-credit-card",
    label: "Bank and Credit Card Statements Request",
  },
  {
    id: "other-statements",
    label:
      "Bank, Credit Card And Other Statements Request (Loans, Other Institutions, Leases)",
  },
] as const;

type Bd1SubMenuId = (typeof BD1_SUB_MENUS)[number]["id"];

type SharePointStatementKind = "bank" | "credit-card";

function classifySharePointActionItem(
  label: string,
): SharePointStatementKind | null {
  const normalized = label.trim().toLowerCase();
  if (!normalized) {
    return null;
  }
  if (
    normalized.includes("credit card statement") ||
    (normalized.includes("credit card") && normalized.includes("statement"))
  ) {
    return "credit-card";
  }
  if (
    normalized.includes("bank statement") ||
    (normalized.includes("bank") && normalized.includes("statement"))
  ) {
    return "bank";
  }
  return null;
}

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

/** Schedule list ids that belong to a BD1/BD5 request type. */
function requestTypeScheduleIds(subMenuId: Bd1SubMenuId): string[] {
  return subMenuId === "other-statements"
    ? ["bd1-other-statements", "bd5-other-statements"]
    : ["bd1", "bd5"];
}

/** BD8 Question List + BD9/BD10 review share the same action items. */
const BD8_QUESTION_SCHEDULE_IDS = ["bd8", "bd9", "bd10"] as const;

function getActionItemsForSchedules(
  store: ClientRequestListsStore,
  clientId: string,
  scheduleIds: readonly string[],
): ActionWorkflowItem[] {
  const allowed = new Set(scheduleIds);
  // Dedupe by label + month so Oct-source and Nov-copy Open rows don't both show.
  const byKey = new Map<string, ActionWorkflowItem>();
  for (const list of store.lists) {
    if (list.clientId !== clientId || !allowed.has(list.scheduleId)) {
      continue;
    }
    for (const item of list.actionItems) {
      const labelKey = item.label.trim().toLowerCase() || item.id;
      const key = `${labelKey}::${item.monthKey}`;
      const existing = byKey.get(key);
      if (!existing) {
        byKey.set(key, item);
        continue;
      }
      if (existing.status === "action" && item.status !== "action") {
        byKey.set(key, item);
      }
    }
  }
  return Array.from(byKey.values());
}

function getActionItemsForRequestType(
  store: ClientRequestListsStore,
  clientId: string,
  subMenuId: Bd1SubMenuId,
): ActionWorkflowItem[] {
  return getActionItemsForSchedules(
    store,
    clientId,
    requestTypeScheduleIds(subMenuId),
  );
}

function bd1EmailTemplateId(subMenuId: Bd1SubMenuId): string {
  return subMenuId === "other-statements" ? "template-2" : "template-1";
}

function findEmailTemplateById(
  store: EmailTemplatesStore,
  templateId: string,
): EmailTemplate | null {
  return (
    store.templates.find((template) => template.id === templateId) ?? null
  );
}

function findBd1EmailTemplate(
  store: EmailTemplatesStore,
  subMenuId: Bd1SubMenuId,
): EmailTemplate | null {
  return findEmailTemplateById(store, bd1EmailTemplateId(subMenuId));
}

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
    case "bd1Template": {
      const template = normalizeBd1Template(profile?.bd1Template ?? "");
      return (
        BD1_TEMPLATE_OPTIONS.find((option) => option.id === template)?.label ??
        "Template 1"
      );
    }
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

export function ClientManagementPanel() {
  const { instance, accounts } = useMsal();
  const { email, canEdit } = useAccess();
  const msalAccount = accounts[0] ?? null;
  const msalName = msalAccount?.name ?? "";
  const emailFrom =
    getAccountEmail(msalAccount) || email.trim().toLowerCase() || "";

  const [resources, setResources] = useState<ResourceClientEntry[]>([]);
  const [profiles, setProfiles] = useState<ClientProfileEntry[]>([]);
  const [portalUsers, setPortalUsers] = useState<PortalUser[]>([]);
  const [portalName, setPortalName] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedScheduleId, setSelectedScheduleId] = useState<string>(
    CLIENT_SCHEDULE_ITEMS[0].id,
  );
  const [lastBdAudience, setLastBdAudience] =
    useState<LastBdAudience>("client");
  const [requestLists, setRequestLists] = useState<ClientRequestListsStore>(
    createEmptyRequestListsStore(),
  );
  const [scheduleSettings, setScheduleSettings] =
    useState<ClientScheduleSettingsStore>(createEmptyScheduleSettingsStore);
  const [inactiveExpanded, setInactiveExpanded] = useState(false);
  const [savingScheduleSettings, setSavingScheduleSettings] = useState(false);
  const [holidays, setHolidays] = useState<HolidaysStore>(createDefaultHolidays);
  const [emailTemplates, setEmailTemplates] = useState<EmailTemplatesStore>(
    createDefaultEmailTemplates,
  );
  const [emailSignature, setEmailSignature] = useState<EmailSignature>(
    createEmptyEmailSignature,
  );
  const [workflowTab, setWorkflowTab] = useState<WorkflowTab>("request");
  const [requestTypeId, setRequestTypeId] =
    useState<Bd1SubMenuId>("bank-credit-card");
  const [draftRequestLabel, setDraftRequestLabel] = useState("");
  const [editingRequestItemId, setEditingRequestItemId] = useState<
    string | null
  >(null);
  const [editRequestLabel, setEditRequestLabel] = useState("");
  const [savingRequestList, setSavingRequestList] = useState(false);
  const [questionListAttachment, setQuestionListAttachment] =
    useState<GraphMailAttachment | null>(null);
  const [uploadingQuestionList, setUploadingQuestionList] = useState(false);
  const questionListFileRef = useRef<HTMLInputElement | null>(null);
  const [bd15Attachments, setBd15Attachments] = useState<
    Bd15StoredAttachment[]
  >([]);
  const [uploadingBd15Files, setUploadingBd15Files] = useState(false);
  const bd15FileRef = useRef<HTMLInputElement | null>(null);
  const [bd15EmailStatus, setBd15EmailStatus] = useState<{
    emailSent: boolean;
    emailSentAt: string;
    comments: Array<{ id: string; text: string; createdAt: string }>;
  }>({ emailSent: false, emailSentAt: "", comments: [] });
  const [bd18CommentDraft, setBd18CommentDraft] = useState("");
  const [savingBd18Comment, setSavingBd18Comment] = useState(false);
  const [savingBd1Template, setSavingBd1Template] = useState(false);
  const [checkingSharePoint, setCheckingSharePoint] = useState(false);
  const [sharePointCheckMessage, setSharePointCheckMessage] = useState<
    string | null
  >(null);
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
  const [sendingEmail, setSendingEmail] = useState(false);
  const [schedulingEmail, setSchedulingEmail] = useState(false);
  const [emailSendMessage, setEmailSendMessage] = useState<string | null>(null);
  const [emailSendError, setEmailSendError] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);
  /** Calendar clock for BD dates — may follow Repeated Tasks demo month. */
  const [scheduleNow, setScheduleNow] = useState(() => new Date());
  const [demoMonthLabel, setDemoMonthLabel] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;

    function applyDemoMonth(demoMonth: string | null | undefined) {
      const clock = getClientManagementNow(demoMonth);
      const nextKey = currentMonthKey(clock);
      setScheduleNow((prev) =>
        currentMonthKey(prev) === nextKey ? prev : clock,
      );
      if (demoMonth && /^\d{4}-\d{2}$/.test(demoMonth)) {
        const [year, mon] = demoMonth.split("-").map(Number);
        setDemoMonthLabel(
          new Date(year, mon - 1, 1).toLocaleString("en-US", {
            month: "long",
            year: "numeric",
          }),
        );
      } else {
        setDemoMonthLabel(null);
      }
    }

    async function loadDemoMonth() {
      try {
        const response = await portalFetch("/api/repeated-tasks", {
          cache: "no-store",
        });
        if (!response.ok || cancelled) {
          return;
        }
        const data = (await response.json()) as RepeatedTasksStore;
        const task = data.tasks?.find(
          (item) => item.id === ROLL_CLIENT_MANAGEMENT_MONTH_ID,
        );
        if (!cancelled) {
          applyDemoMonth(task?.demoMonth ?? null);
        }
      } catch {
        // Keep real calendar if repeated-tasks is unavailable.
      }
    }

    void loadDemoMonth();
    const onDemoMonth = () => {
      void loadDemoMonth();
    };
    window.addEventListener(CLIENT_MANAGEMENT_DEMO_MONTH_EVENT, onDemoMonth);
    return () => {
      cancelled = true;
      window.removeEventListener(
        CLIENT_MANAGEMENT_DEMO_MONTH_EVENT,
        onDemoMonth,
      );
    };
  }, []);

  const showBd1TypeContent = selectedScheduleId === "bd1";
  const showBd5Content = selectedScheduleId === "bd5";
  const showBd7Content = selectedScheduleId === "bd7";
  const showBd8Content = selectedScheduleId === "bd8";
  const showBd15Content = selectedScheduleId === "bd15";
  /** BD18 is the status check for BD15 books-closed email/files. */
  const showBd18Content = selectedScheduleId === "bd18";
  /** Last 3 / 2 / 1 BD sidebar days: Client | Team email only. */
  const showLastBdSchedule = isLastBdSchedule(selectedScheduleId);
  /** Last 3 BD → Client: Template 5. */
  const showLast3BdContent =
    selectedScheduleId === "last-3-bd" && lastBdAudience === "client";
  /** Last 2 BD → Client: Template 12. */
  const showLast2BdContent =
    selectedScheduleId === "last-2-bd" && lastBdAudience === "client";
  /** Last 1 BD → Client: end-of-month follow-up (no numbered template yet). */
  const showLast1BdClientContent =
    selectedScheduleId === "last-1-bd" && lastBdAudience === "client";
  /** Last N BD → Team: Template 18 to portal users. */
  const showLastBdTeamContent =
    showLastBdSchedule && lastBdAudience === "team";
  const showEndOfMonthEmailOnly = showLastBdSchedule;
  /** BD9/BD10 review BD8 questions (Completed/Late) + reminder email. */
  const showBd9Content = selectedScheduleId === "bd9";
  const showBd10Content = selectedScheduleId === "bd10";
  const showBd9LikeContent = showBd9Content || showBd10Content;
  const showReviewActionTable = showBd5Content || showBd9LikeContent;
  /** BD1/BD8/BD15/BD18 share Request/Attachments + Email (no Action Items tab). */
  const showBd1LikeRequestList =
    showBd1TypeContent ||
    showBd8Content ||
    showBd15Content ||
    showBd18Content;
  /** BD15 and BD18 share the same Template 11 email + attachments. */
  const showBd15Email = showBd15Content || showBd18Content;
  const showBd1LikeWorkflow =
    showBd1TypeContent ||
    showBd7Content ||
    showBd8Content ||
    showBd9LikeContent ||
    showBd15Email ||
    showEndOfMonthEmailOnly;
  const showRequestTypeChooser =
    selectedScheduleId === "bd1" ||
    selectedScheduleId === "bd5" ||
    selectedScheduleId === "bd7";
  const showRequestListTab = showBd1LikeRequestList;
  const activeRequestType =
    BD1_SUB_MENUS.find((item) => item.id === requestTypeId) ??
    BD1_SUB_MENUS[0];
  const activeRequestScheduleId = showRequestTypeChooser
    ? requestTypeScheduleId(selectedScheduleId, requestTypeId)
    : selectedScheduleId;

  useEffect(() => {
    setEditingRequestItemId(null);
    setEditRequestLabel("");
    setSharePointCheckMessage(null);
  }, [selectedId, selectedScheduleId, workflowTab, requestTypeId]);

  // Load saved BD8 Excel attachment for this client + current month.
  // Prior-month files are pruned server-side once when the month changes.
  useEffect(() => {
    if (!selectedId || selectedScheduleId !== "bd8") {
      setQuestionListAttachment(null);
      if (questionListFileRef.current) {
        questionListFileRef.current.value = "";
      }
      return;
    }

    let cancelled = false;
    void (async () => {
      try {
        const response = await portalFetch(
          `/api/bd8-email-attachments?clientId=${encodeURIComponent(selectedId)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}`,
          { cache: "no-store" },
        );
        const data = (await response.json()) as {
          attachment?: GraphMailAttachment & {
            clientId?: string;
            monthKey?: string;
          } | null;
          error?: string;
        };
        if (cancelled) {
          return;
        }
        if (!response.ok) {
          throw new Error(data.error || "Failed to load Excel attachment.");
        }
        const saved = data.attachment;
        if (
          saved &&
          typeof saved.name === "string" &&
          typeof saved.contentBytes === "string" &&
          saved.name &&
          saved.contentBytes
        ) {
          setQuestionListAttachment({
            name: saved.name,
            contentType: saved.contentType || "application/octet-stream",
            contentBytes: saved.contentBytes,
          });
        } else {
          setQuestionListAttachment(null);
          if (questionListFileRef.current) {
            questionListFileRef.current.value = "";
          }
        }
      } catch {
        if (!cancelled) {
          setQuestionListAttachment(null);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [selectedId, selectedScheduleId, scheduleNow]);

  // Load saved BD15 PDF attachments for this client + current month.
  // Also load on BD18 so status can show attached / not attached.
  useEffect(() => {
    if (
      !selectedId ||
      (selectedScheduleId !== "bd15" && selectedScheduleId !== "bd18")
    ) {
      setBd15Attachments([]);
      if (bd15FileRef.current) {
        bd15FileRef.current.value = "";
      }
      return;
    }

    let cancelled = false;
    void (async () => {
      try {
        const response = await portalFetch(
          `/api/bd15-email-attachments?clientId=${encodeURIComponent(selectedId)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}`,
          { cache: "no-store" },
        );
        const data = (await response.json()) as {
          attachments?: Bd15ApiAttachment[];
          error?: string;
        };
        if (cancelled) {
          return;
        }
        if (!response.ok) {
          throw new Error(data.error || "Failed to load PDF attachments.");
        }
        setBd15Attachments(mapBd15ApiAttachments(data.attachments));
      } catch {
        if (!cancelled) {
          setBd15Attachments([]);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [selectedId, selectedScheduleId, scheduleNow]);

  // Load BD15 email-sent status + comments (used by BD18 status check).
  useEffect(() => {
    if (
      !selectedId ||
      (selectedScheduleId !== "bd15" && selectedScheduleId !== "bd18")
    ) {
      setBd15EmailStatus({
        emailSent: false,
        emailSentAt: "",
        comments: [],
      });
      setBd18CommentDraft("");
      return;
    }

    let cancelled = false;
    void (async () => {
      try {
        const response = await portalFetch(
          `/api/bd15-email-status?clientId=${encodeURIComponent(selectedId)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}`,
          { cache: "no-store" },
        );
        const data = (await response.json()) as {
          status?: {
            emailSent?: boolean;
            emailSentAt?: string;
            comments?: Array<{
              id?: string;
              text?: string;
              createdAt?: string;
            }>;
          };
          error?: string;
        };
        if (cancelled) {
          return;
        }
        if (!response.ok) {
          throw new Error(data.error || "Failed to load BD15 email status.");
        }
        const status = data.status;
        setBd15EmailStatus({
          emailSent: Boolean(status?.emailSent),
          emailSentAt:
            typeof status?.emailSentAt === "string" ? status.emailSentAt : "",
          comments: Array.isArray(status?.comments)
            ? status.comments
                .filter(
                  (item) =>
                    typeof item.id === "string" &&
                    item.id &&
                    typeof item.text === "string" &&
                    item.text.trim(),
                )
                .map((item) => ({
                  id: item.id as string,
                  text: (item.text as string).trim(),
                  createdAt:
                    typeof item.createdAt === "string" ? item.createdAt : "",
                }))
            : [],
        });
      } catch {
        if (!cancelled) {
          setBd15EmailStatus({
            emailSent: false,
            emailSentAt: "",
            comments: [],
          });
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [selectedId, selectedScheduleId, scheduleNow]);

  useEffect(() => {
    if (!showRequestListTab && workflowTab === "request") {
      setWorkflowTab(showBd7Content ? "action" : showBd1LikeWorkflow ? "email" : "action");
    }
  }, [showRequestListTab, showBd7Content, showBd1LikeWorkflow, workflowTab]);

  useEffect(() => {
    // BD1/BD8/BD15/BD18 no longer show Action Items — keep users on
    // Request/Attachments / Email.
    if (showBd1LikeRequestList && workflowTab === "action") {
      setWorkflowTab("request");
    }
  }, [showBd1LikeRequestList, workflowTab]);

  useEffect(() => {
    if (workflowTab === "email" && !showBd1LikeWorkflow) {
      setWorkflowTab("action");
    }
  }, [workflowTab, showBd1LikeWorkflow]);

  useEffect(() => {
    if (showBd1LikeWorkflow && workflowTab === "completed") {
      setWorkflowTab(
        showBd1LikeRequestList
          ? "request"
          : showEndOfMonthEmailOnly
            ? "email"
            : "action",
      );
    }
  }, [
    showBd1LikeWorkflow,
    showBd1LikeRequestList,
    showEndOfMonthEmailOnly,
    workflowTab,
  ]);

  useEffect(() => {
    if (showEndOfMonthEmailOnly && workflowTab !== "email") {
      setWorkflowTab("email");
    }
  }, [showEndOfMonthEmailOnly, workflowTab]);

  useEffect(() => {
    if (showRequestTypeChooser) {
      return;
    }
    setRequestTypeId("bank-credit-card");
  }, [showRequestTypeChooser, selectedScheduleId]);

  // BD5 is status-only (no tabs). BD9/BD10 have Question List Status + Email.
  useEffect(() => {
    if (showBd5Content && workflowTab !== "action") {
      setWorkflowTab("action");
    }
  }, [showBd5Content, workflowTab]);

  useEffect(() => {
    if (
      showBd9LikeContent &&
      workflowTab !== "action" &&
      workflowTab !== "email"
    ) {
      setWorkflowTab("action");
    }
  }, [showBd9LikeContent, workflowTab]);

  const load = useCallback(async () => {
    setError(null);
    try {
      const [
        resourcesRes,
        profilesRes,
        usersRes,
        requestListsRes,
        scheduleSettingsRes,
        holidaysRes,
        emailTemplatesRes,
        signatureRes,
      ] = await Promise.all([
        fetch("/api/resource-clients", { cache: "no-store" }),
        fetch("/api/client-profiles", { cache: "no-store" }),
        fetch("/api/portal-users", { cache: "no-store" }),
        fetch("/api/client-request-lists", { cache: "no-store" }),
        fetch("/api/client-schedule-settings", { cache: "no-store" }),
        fetch("/api/holidays", { cache: "no-store" }),
        fetch("/api/email-templates", { cache: "no-store" }),
        portalFetch("/api/email-signature", { cache: "no-store" }),
      ]);
      const resourcesData = (await resourcesRes.json()) as ResourcesResponse;
      const profilesData = (await profilesRes.json()) as ProfilesResponse;
      const usersData = (await usersRes.json()) as UsersResponse;
      const requestListsData = await requestListsRes.json();
      const scheduleSettingsData = await scheduleSettingsRes.json();
      const holidaysData = await holidaysRes.json();
      const emailTemplatesData = await emailTemplatesRes.json();
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
      setPortalUsers(Array.isArray(usersData.users) ? usersData.users : []);
      setPortalName(matchedUser?.name?.trim() || "");
      // Keep existing in-memory lists if the API fails — do not wipe UI to
      // empty (that previously got saved over a corrupt/empty data file).
      if (requestListsRes.ok) {
        setRequestLists(normalizeClientRequestListsStore(requestListsData));
      }
      setScheduleSettings(
        scheduleSettingsRes.ok
          ? normalizeClientScheduleSettingsStore(scheduleSettingsData)
          : createEmptyScheduleSettingsStore(),
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

  // When demo/real month changes, reload lists and clear in-progress form edits.
  const scheduleMonthKeyForLists = currentMonthKey(scheduleNow);
  useEffect(() => {
    let cancelled = false;
    setDraftRequestLabel("");
    setEditingRequestItemId(null);
    setEditRequestLabel("");
    void (async () => {
      try {
        const response = await fetch(
          `/api/client-request-lists?monthKey=${encodeURIComponent(scheduleMonthKeyForLists)}`,
          { cache: "no-store" },
        );
        if (!response.ok || cancelled) {
          return;
        }
        const data = await response.json();
        if (!cancelled) {
          setRequestLists(normalizeClientRequestListsStore(data));
        }
      } catch {
        // Keep existing lists if reload fails.
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [scheduleMonthKeyForLists]);

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
  const selectedBd1Template = normalizeBd1Template(
    selected?.profile?.bd1Template ?? "template-1",
  );
  const activeScheduleIdSet = useMemo(() => {
    if (!selectedClientId) {
      return new Set<string>();
    }
    return new Set(
      getActiveScheduleIds(
        scheduleSettings,
        selectedClientId,
        ALL_SCHEDULE_IDS,
      ),
    );
  }, [scheduleSettings, selectedClientId]);
  const sortScheduleItemsByMonthDate = useCallback(
    <T extends { id: string }>(items: readonly T[]): T[] => {
      const country = selected?.profile?.country;
      const state = selected?.profile?.state ?? "";
      return [...items].sort((a, b) => {
        const dateA = getScheduleBusinessDate(
          a.id,
          holidays,
          country,
          state,
          scheduleNow,
        );
        const dateB = getScheduleBusinessDate(
          b.id,
          holidays,
          country,
          state,
          scheduleNow,
        );
        const timeA = dateA?.getTime() ?? Number.MAX_SAFE_INTEGER;
        const timeB = dateB?.getTime() ?? Number.MAX_SAFE_INTEGER;
        if (timeA !== timeB) {
          return timeA - timeB;
        }
        return (
          (CLIENT_SCHEDULE_ORDER.get(a.id) ?? 0) -
          (CLIENT_SCHEDULE_ORDER.get(b.id) ?? 0)
        );
      });
    },
    [selected?.profile?.country, selected?.profile?.state, holidays, scheduleNow],
  );

  const activeScheduleItems = useMemo(
    () =>
      sortScheduleItemsByMonthDate(
        CLIENT_SCHEDULE_ITEMS.filter((item) =>
          activeScheduleIdSet.has(item.id),
        ),
      ),
    [activeScheduleIdSet, sortScheduleItemsByMonthDate],
  );
  const inactiveScheduleItems = useMemo(
    () =>
      sortScheduleItemsByMonthDate(
        CLIENT_SCHEDULE_ITEMS.filter(
          (item) => !activeScheduleIdSet.has(item.id),
        ),
      ),
    [activeScheduleIdSet, sortScheduleItemsByMonthDate],
  );
  const isSelectedScheduleActive = Boolean(
    selectedClientId &&
      isScheduleActiveForClient(
        scheduleSettings,
        selectedClientId,
        selectedScheduleId,
        ALL_SCHEDULE_IDS,
      ),
  );

  // When switching clients, land on an Active business day if the current one is not.
  useEffect(() => {
    if (!selectedClientId) {
      return;
    }
    const activeIds = getActiveScheduleIds(
      scheduleSettings,
      selectedClientId,
      ALL_SCHEDULE_IDS,
    );
    setSelectedScheduleId((current) =>
      activeIds.includes(current)
        ? current
        : (activeIds[0] ?? CLIENT_SCHEDULE_ITEMS[0].id),
    );
    setInactiveExpanded(false);
    // Only re-run when the selected client changes.
    // eslint-disable-next-line react-hooks/exhaustive-deps -- intentional
  }, [selectedClientId]);

  // BD1/BD5/BD7 request type follows the client's BD1 Template (default Template 1).
  useEffect(() => {
    if (!showRequestTypeChooser) {
      return;
    }
    setRequestTypeId(bd1TemplateToRequestTypeId(selectedBd1Template));
  }, [showRequestTypeChooser, selectedBd1Template, selectedClientId]);

  const defaultEmailDeadline = useMemo(() => {
    if (
      !selectedClientId ||
      (selectedScheduleId !== "bd1" &&
        selectedScheduleId !== "bd7" &&
        selectedScheduleId !== "bd8" &&
        selectedScheduleId !== "bd9" &&
        selectedScheduleId !== "bd10")
    ) {
      return "";
    }
    // BD1 → BD4; BD7 → BD8 (or BD7 if weekend between); BD8/9/10 → same day.
    const dueDate =
      selectedScheduleId === "bd7"
        ? getBd7DeadlineDate(
            holidays,
            selected?.profile?.country,
            selected?.profile?.state ?? "",
            scheduleNow,
          )
        : getScheduleBusinessDate(
            selectedScheduleId === "bd1" ? "bd4" : selectedScheduleId,
            holidays,
            selected?.profile?.country,
            selected?.profile?.state ?? "",
            scheduleNow,
          );
    if (!dueDate) {
      return "";
    }
    const year = dueDate.getFullYear();
    const month = String(dueDate.getMonth() + 1).padStart(2, "0");
    const day = String(dueDate.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }, [
    selectedClientId,
    selectedScheduleId,
    selected?.profile?.country,
    selected?.profile?.state,
    holidays,
    scheduleNow,
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
    // Migrate legacy Last N BD (Team) sidebar ids into the merged slot + Team toggle.
    const legacyTeam = /^last-(\d+)-bd-team$/i.exec(selectedScheduleId.trim());
    if (legacyTeam) {
      setSelectedScheduleId(`last-${legacyTeam[1]}-bd`);
      setLastBdAudience("team");
      return;
    }
  }, [selectedScheduleId]);

  useEffect(() => {
    if (
      !selectedClientId ||
      (selectedScheduleId !== "bd1" &&
        selectedScheduleId !== "bd7" &&
        selectedScheduleId !== "bd8" &&
        selectedScheduleId !== "bd9" &&
        selectedScheduleId !== "bd10" &&
        selectedScheduleId !== "bd15" &&
        selectedScheduleId !== "bd18" &&
        !showLast3BdContent &&
        !showLast2BdContent &&
        !showLast1BdClientContent &&
        !showLastBdTeamContent)
    ) {
      return;
    }
    if (showLast1BdClientContent) {
      setEmailSubject("");
      setEmailBody("");
      setEmailPortalHref("");
      setEmailAutofillHighlights([]);
      return;
    }
    const template = showLast3BdContent
      ? findEmailTemplateById(emailTemplates, "template-5")
      : showLast2BdContent
        ? findEmailTemplateById(emailTemplates, "template-12")
        : showLastBdTeamContent
          ? findEmailTemplateById(emailTemplates, "template-18")
          : selectedScheduleId === "bd15" || selectedScheduleId === "bd18"
            ? findEmailTemplateById(emailTemplates, "template-11")
            : selectedScheduleId === "bd10"
              ? findEmailTemplateById(emailTemplates, "template-7")
              : selectedScheduleId === "bd9"
                ? findEmailTemplateById(emailTemplates, "template-8")
                : selectedScheduleId === "bd8"
                  ? findEmailTemplateById(emailTemplates, "template-3")
                  : findBd1EmailTemplate(emailTemplates, requestTypeId);
    const clientName =
      selectedOperatingName.replace(/[.:]+$/g, "").trim() ||
      selectedOperatingName.trim() ||
      "[Client Name]";
    const contactName = selectedPointOfContact || "[Contact First Name]";
    const portalLink = selectedPortalLink || "[Portal Link]";
    const periodDate = scheduleNow;
    const period = formatCurrentPeriod(periodDate);
    const periodPortalHref = resolvePortalHref(portalLink, periodDate);
    // Template 5 (Last 3 BD): formal request is next month's BD1 / due BD4.
    const nextMonthAnchor = new Date(
      scheduleNow.getFullYear(),
      scheduleNow.getMonth() + 1,
      15,
    );
    const nextBd1Date = getScheduleBusinessDate(
      "bd1",
      holidays,
      selected?.profile?.country,
      selected?.profile?.state ?? "",
      nextMonthAnchor,
    );
    const nextBd4Date = getScheduleBusinessDate(
      "bd4",
      holidays,
      selected?.profile?.country,
      selected?.profile?.state ?? "",
      nextMonthAnchor,
    );
    const bd1DateLabel = nextBd1Date
      ? formatBusinessDayLabel(nextBd1Date)
      : "[BD1 Date]";
    const dueDate = showLast3BdContent
      ? nextBd4Date
        ? formatBusinessDayLabel(nextBd4Date)
        : emailDeadlineLabel || "[Due Date]"
      : emailDeadlineLabel;
    const emailMonthKey = currentMonthKey(periodDate);
    const typeActionItems = getActionItemsForRequestType(
      requestLists,
      selectedClientId,
      requestTypeId,
    ).filter((item) => item.monthKey === emailMonthKey);
    const bd8QuestionActionItems = getActionItemsForSchedules(
      requestLists,
      selectedClientId,
      BD8_QUESTION_SCHEDULE_IDS,
    ).filter((item) => item.monthKey === emailMonthKey);
    const requestListScheduleId =
      selectedScheduleId === "bd8"
        ? "bd8"
        : requestTypeScheduleId("bd1", requestTypeId);
    const requestListItems = requestItemsForMonth(
      getClientRequestList(
        requestLists,
        selectedClientId,
        requestListScheduleId,
      ).requestItems,
      emailMonthKey,
    );
    const last3BdStatementItems = showLast3BdContent
      ? [
          ...requestItemsForMonth(
            getClientRequestList(requestLists, selectedClientId, "bd1")
              .requestItems,
            emailMonthKey,
          ),
          ...requestItemsForMonth(
            getClientRequestList(
              requestLists,
              selectedClientId,
              "bd1-other-statements",
            ).requestItems,
            emailMonthKey,
          ),
        ]
      : [];
    // BD1/BD8 email uses Request List Template bullets directly.
    // BD7 email still uses items marked Late on BD5.
    // BD9 reminder uses open + late BD8 questions; BD10 uses only Late from BD9.
    // Last 3 BD Client (Template 5) uses BD1 request-list bullets as the statements list.
    const emailListItems = showLast3BdContent
      ? last3BdStatementItems
      : selectedScheduleId === "bd7"
        ? typeActionItems.filter((item) => item.status === "late")
        : selectedScheduleId === "bd10"
          ? bd8QuestionActionItems.filter((item) => item.status === "late")
          : selectedScheduleId === "bd9"
            ? bd8QuestionActionItems.filter(
                (item) =>
                  item.status === "action" || item.status === "late",
              )
            : requestListItems;
    const bankListPlaceholder =
      "• [List of bank and credit card accounts, with last 4 digits]";
    const otherListPlaceholder =
      "• [List of loan, line of credit, lease, investment or other accounts]";
    const outstandingListPlaceholder = "• [List of outstanding items only]";
    const statementsListPlaceholder =
      "• [List of statements and reports for this client]";
    const actionListText =
      emailListItems.length > 0
        ? emailListItems.map((item) => `• ${item.label}`).join("\n")
        : showLast3BdContent
          ? statementsListPlaceholder
          : selectedScheduleId === "bd9" || selectedScheduleId === "bd10"
            ? outstandingListPlaceholder
            : requestTypeId === "other-statements"
              ? otherListPlaceholder
              : bankListPlaceholder;
    const openQuestionCount = showLast2BdContent
      ? bd8QuestionActionItems.filter(
          (item) =>
            item.status === "action" || item.status === "late",
        ).length
      : emailListItems.length;
    const questionCount =
      openQuestionCount > 0 ? String(openQuestionCount) : "[Number]";
    const questionListText =
      emailListItems.length > 0
        ? emailListItems.map((item) => `• ${item.label}`).join("\n")
        : "• [Question List]";

    function fillPlaceholders(text: string) {
      return applyPeriodPlaceholder(
        applyPortalLinkPlaceholders(
          text
            .replace(/\[Client Name\]/gi, clientName)
            .replace(/\[Contact First Name\]/gi, contactName)
            .replace(/\[Due Date\]/gi, dueDate)
            .replace(/\[Deadline\]/gi, dueDate)
            .replace(/\[BD1 Date\]/gi, bd1DateLabel)
            .replace(
              /\[Last BD Label\]/gi,
              lastBdLabel(selectedScheduleId),
            )
            .replace(/\[Number\]/gi, questionCount)
            .replace(/(?:•\s*)?\[Question List\]/gi, questionListText)
            .replace(
              /(?:•\s*)?\[List of statements and reports for this client\]/gi,
              actionListText,
            )
            .replace(
              /(?:•\s*)?\[List of outstanding items only\]/gi,
              actionListText,
            )
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
        periodDate,
      );
    }

    const subject = fillPlaceholders(template?.subject?.trim() || "");
    let body = fillPlaceholders(
      stripLegacySignatureClosing(template?.body ?? "").trimEnd(),
    );
    // BD8/BD9/BD10: ensure the question list bullets appear in the email body
    // (same idea as Template 7's outstanding-items list on BD10).
    const isQuestionEmailDay =
      selectedScheduleId === "bd8" ||
      selectedScheduleId === "bd9" ||
      selectedScheduleId === "bd10";
    if (
      isQuestionEmailDay &&
      emailListItems.length > 0 &&
      !emailListItems.every((item) => body.includes(item.label))
    ) {
      const withList = body
        .replace(
          /You can see the full list in the portal(?:\s+here)?:?\s*(?:\[Portal Link\])?\s*(?:\[Question List\])?/i,
          `You can see the full list in the portal:\n\n${questionListText}\n\n`,
        )
        .replace(
          /You can see and answer them(?:\s+here)?:?\s*(?:\[Portal Link\])?\s*(?:\[Question List\])?/i,
          `You can see and answer them here:\n\n${questionListText}\n\n`,
        )
        .replace(
          /(?:•\s*)?\[List of outstanding items only\]/i,
          questionListText,
        )
        .replace(/(?:•\s*)?\[Question List\]/i, questionListText);
      if (withList !== body) {
        body = withList;
      } else {
        body = `${body.trimEnd()}\n\n${questionListText}`;
      }
    }
    const lastBdLabelText = lastBdLabel(selectedScheduleId);
    const highlights = [
      clientName,
      contactName,
      period,
      dueDate,
      bd1DateLabel !== "[BD1 Date]" ? bd1DateLabel : "",
      lastBdLabelText !== "[Last BD Label]" ? lastBdLabelText : "",
      questionCount !== "[Number]" ? questionCount : "",
      ...emailListItems.map((item) => item.label),
    ].filter(
      (value) =>
        value &&
        value !== "[Client Name]" &&
        value !== "[Contact First Name]" &&
        value !== "[Due Date]" &&
        value !== "[Deadline]" &&
        value !== "[Portal Link]" &&
        value !== "[Number]",
    );

    setEmailSubject(subject);
    setEmailBody(body);
    setEmailPortalHref(periodPortalHref);
    setEmailAutofillHighlights(highlights);
  }, [
    selectedClientId,
    selectedScheduleId,
    showLast3BdContent,
    showLast2BdContent,
    showLast1BdClientContent,
    showLastBdTeamContent,
    requestTypeId,
    selectedOperatingName,
    selectedPointOfContact,
    selectedPortalLink,
    emailDeadlineLabel,
    requestLists,
    emailTemplates,
    scheduleNow,
    holidays,
    selected?.profile?.country,
    selected?.profile?.state,
  ]);

  useEffect(() => {
    if (
      !selectedClientId ||
      (selectedScheduleId !== "bd1" &&
        selectedScheduleId !== "bd7" &&
        selectedScheduleId !== "bd8" &&
        selectedScheduleId !== "bd9" &&
        selectedScheduleId !== "bd10" &&
        selectedScheduleId !== "bd15" &&
        selectedScheduleId !== "bd18" &&
        !showLast3BdContent &&
        !showLast2BdContent &&
        !showLast1BdClientContent &&
        !showLastBdTeamContent)
    ) {
      return;
    }
    // Leave To empty on load; From always follows the signed-in account.
    // Team end-of-month: To = sender, BCC = all portal users.
    if (showLastBdTeamContent) {
      const teamEmails = portalUsers
        .map((user) => user.email.trim().toLowerCase())
        .filter(Boolean);
      const sender = emailFrom.trim().toLowerCase();
      const bccList = teamEmails.filter((item) => item !== sender);
      setEmailTo(sender || teamEmails[0] || "");
      setEmailCc("");
      setEmailBcc(bccList.join(", "));
    } else {
      setEmailTo("");
      setEmailCc("");
      setEmailBcc("");
    }
    setEmailDeadline("");
    setEmailSendAt("");
    setEmailSendError(null);
    setEmailSendMessage(null);
  }, [
    selectedClientId,
    selectedScheduleId,
    showLast3BdContent,
    showLast2BdContent,
    showLast1BdClientContent,
    showLastBdTeamContent,
    requestTypeId,
    portalUsers,
    emailFrom,
  ]);

  const selectedSchedule =
    CLIENT_SCHEDULE_ITEMS.find((item) => item.id === selectedScheduleId) ??
    CLIENT_SCHEDULE_ITEMS[0];

  const selectedScheduleBusinessDate = useMemo(() => {
    if (!selected) {
      return null;
    }
    return getScheduleBusinessDate(
      selectedScheduleId,
      holidays,
      selected.profile?.country,
      selected.profile?.state ?? "",
      scheduleNow,
    );
  }, [selected, selectedScheduleId, holidays, scheduleNow]);

  // BD date at 9am client TZ → Pakistan (default; editable override in state).
  const clientTimeZoneMissing = !selected?.profile?.timeZone?.trim();
  const clientTimeZoneInvalid =
    !clientTimeZoneMissing &&
    !resolveClientTimeZone(selected?.profile?.timeZone);
  const clientTimeZoneError = clientTimeZoneMissing
    ? CLIENT_TIME_ZONE_REQUIRED_MESSAGE
    : clientTimeZoneInvalid
      ? "Client time zone is not recognized. Set a valid time zone in Client Profile."
      : null;
  const defaultEmailSendSchedule = useMemo(
    () =>
      getClientNineAmSchedule(
        selectedScheduleBusinessDate,
        selected?.profile?.timeZone,
      ),
    [selectedScheduleBusinessDate, selected?.profile?.timeZone],
  );
  const emailSendAtClientLabel = defaultEmailSendSchedule?.clientLabel ?? "";
  const effectiveEmailSendAt =
    emailSendAt || defaultEmailSendSchedule?.userDateTimeLocal || "";

  const scheduleDateById = useMemo(() => {
    const map = new Map<string, string>();
    if (!selected) {
      return map;
    }
    const country = selected.profile?.country;
    const state = selected.profile?.state ?? "";
    for (const item of CLIENT_SCHEDULE_ITEMS) {
      const date = getScheduleBusinessDate(
        item.id,
        holidays,
        country,
        state,
        scheduleNow,
      );
      if (date) {
        map.set(item.id, formatBusinessDayLabel(date));
      }
    }
    return map;
  }, [selected, holidays, scheduleNow]);

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

  const scheduleMonthKey = currentMonthKey(scheduleNow);
  /** BD1/BD8 request bullets for the active (demo) month only. */
  const visibleRequestItems = requestItemsForMonth(
    activeList.requestItems,
    scheduleMonthKey,
  );
  const sharedActionItems = selected
    ? getClientActionItems(requestLists, selected.resource.id)
    : [];
  const questionScheduleActionItems =
    showBd9LikeContent && selected
      ? getActionItemsForSchedules(
          requestLists,
          selected.resource.id,
          BD8_QUESTION_SCHEDULE_IDS,
        )
      : [];
  const monthScopedActionItems = (
    items: ActionWorkflowItem[],
  ): ActionWorkflowItem[] =>
    items.filter((item) => item.monthKey === scheduleMonthKey);
  const scopedActionItems = monthScopedActionItems(
    showBd10Content
      ? // BD10 only surfaces items marked Late on BD9 (not Completed).
        questionScheduleActionItems.filter((item) => item.status === "late")
      : showBd9Content
        ? questionScheduleActionItems
        : showRequestTypeChooser && selected
          ? getActionItemsForRequestType(
              requestLists,
              selected.resource.id,
              requestTypeId,
            )
          : sharedActionItems,
  );
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
    if (showBd10Content) {
      return scopedActionItems
        .filter((item) => item.status === "late")
        .sort((a, b) => {
          const aTime = Date.parse(a.completedAt) || 0;
          const bTime = Date.parse(b.completedAt) || 0;
          return bTime - aTime;
        });
    }
    const open = scopedActionItems.filter((item) => item.status === "action");
    const resolved = scopedActionItems
      .filter((item) => item.status === "completed" || item.status === "late")
      .sort((a, b) => {
        const aTime = Date.parse(a.completedAt) || 0;
        const bTime = Date.parse(b.completedAt) || 0;
        return bTime - aTime;
      });
    return [...open, ...resolved];
  }, [scopedActionItems, showBd10Content]);

  const bd7EmailDisabled =
    showBd7Content && bd7LateItems.length === 0;
  const bd10EmailDisabled =
    showBd10Content && bd5ActionTableItems.length === 0;
  const bd15MissingReports = useMemo(
    () =>
      showBd15Content || showBd18Content
        ? getMissingBd15Reports(bd15Attachments)
        : [],
    [showBd15Content, showBd18Content, bd15Attachments],
  );
  /** BD15 and BD18 share sent status — after send, email is view-only. */
  const bd15EmailAlreadySent =
    showBd15Email && bd15EmailStatus.emailSent;
  const bd15ReportsMissingForSend =
    showBd15Email &&
    !bd15EmailAlreadySent &&
    bd15MissingReports.length > 0;
  /** Tab stays available after send so the email can still be viewed. */
  const reminderEmailTabDisabled =
    bd7EmailDisabled ||
    bd10EmailDisabled ||
    (showBd15Email &&
      !bd15EmailStatus.emailSent &&
      bd15MissingReports.length > 0);
  const reminderEmailSendDisabled =
    reminderEmailTabDisabled || bd15EmailAlreadySent;
  const reminderEmailDisabled = reminderEmailSendDisabled;
  const reminderEmailDisabledTitle = bd15EmailAlreadySent
    ? `Email sent on ${formatActionResolvedAt(bd15EmailStatus.emailSentAt) || "—"}`
    : bd7EmailDisabled
      ? "No late action items for BD5 found."
      : bd10EmailDisabled
        ? "No late BD9 questions found."
        : bd15ReportsMissingForSend
          ? `These files are missing: ${bd15MissingReports.join(", ")}`
          : undefined;

  useEffect(() => {
    if (reminderEmailTabDisabled && workflowTab === "email") {
      setWorkflowTab("request");
    }
  }, [reminderEmailTabDisabled, workflowTab]);

  async function handleToggleScheduleStar(
    scheduleId: string,
    event: MouseEvent,
  ) {
    event.preventDefault();
    event.stopPropagation();
    if (!selected || !canEdit || savingScheduleSettings) {
      return;
    }

    const nextStore = toggleClientScheduleActive(
      scheduleSettings,
      selected.resource.id,
      scheduleId,
      ALL_SCHEDULE_IDS,
    );
    const entry = {
      clientId: selected.resource.id,
      activeScheduleIds: getActiveScheduleIds(
        nextStore,
        selected.resource.id,
        ALL_SCHEDULE_IDS,
      ),
    };

    setSavingScheduleSettings(true);
    setError(null);
    try {
      const response = await portalFetch("/api/client-schedule-settings", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(entry),
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(
          (data as { error?: string }).error ||
            "Failed to update Active business days.",
        );
      }
      setScheduleSettings(normalizeClientScheduleSettingsStore(data));
    } catch (saveError) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to update Active business days.",
      );
    } finally {
      setSavingScheduleSettings(false);
    }
  }

  async function handleBd1TemplateChange(nextTemplate: Bd1TemplateId) {
    if (!selected?.profile || !canEdit) {
      setError("Link a Client Profile to set BD1 Template.");
      return;
    }
    if (normalizeBd1Template(selected.profile.bd1Template) === nextTemplate) {
      return;
    }

    const updated = normalizeClientProfile({
      ...selected.profile,
      bd1Template: nextTemplate,
    });
    const nextProfiles = profiles.map((profile) =>
      profile.id === updated.id ? updated : profile,
    );

    setSavingBd1Template(true);
    setError(null);
    try {
      const response = await portalFetch("/api/client-profiles", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ profiles: nextProfiles }),
      });
      const data = (await response.json()) as ProfilesResponse;
      if (!response.ok) {
        throw new Error(data.error || "Failed to save BD1 Template.");
      }
      setProfiles(
        Array.isArray(data.profiles) ? data.profiles : nextProfiles,
      );
      setRequestTypeId(bd1TemplateToRequestTypeId(nextTemplate));
    } catch (saveError) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save BD1 Template.",
      );
    } finally {
      setSavingBd1Template(false);
    }
  }

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
        body: JSON.stringify({
          ...nextList,
          monthKey: currentMonthKey(scheduleNow),
        }),
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

  function requireActiveSchedule(actionLabel = "perform this action"): boolean {
    if (isSelectedScheduleActive) {
      return true;
    }
    setError(
      `This business day is Inactive. Fill the star to ${actionLabel}.`,
    );
    return false;
  }

  async function handleAddRequestItem() {
    const label = draftRequestLabel.trim();
    if (!label || !selected || !activeRequestScheduleId) {
      return;
    }
    if (!requireActiveSchedule("edit the request list")) {
      return;
    }
    const nextList = addRequestBulletToList(
      {
        ...activeList,
        clientId: selected.resource.id,
        scheduleId: activeRequestScheduleId,
      },
      label,
      scheduleNow,
    );
    setDraftRequestLabel("");
    await persistList(nextList);
  }

  async function handleQuestionListExcelUpload(
    event: ChangeEvent<HTMLInputElement>,
  ) {
    const file = event.target.files?.[0];
    if (!file || !selected) {
      return;
    }
    if (!requireActiveSchedule("attach the Excel file")) {
      event.target.value = "";
      return;
    }

    setUploadingQuestionList(true);
    setError(null);
    try {
      const buffer = await file.arrayBuffer();
      const attachment: GraphMailAttachment = {
        name: file.name,
        contentType: mimeTypeForExcelFile(file),
        contentBytes: arrayBufferToBase64(buffer),
      };
      const response = await portalFetch("/api/bd8-email-attachments", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          clientId: selected.resource.id,
          monthKey: currentMonthKey(scheduleNow),
          ...attachment,
        }),
      });
      const data = (await response.json()) as {
        attachment?: GraphMailAttachment | null;
        error?: string;
      };
      if (!response.ok || !data.attachment) {
        throw new Error(data.error || "Failed to save Excel attachment.");
      }
      setQuestionListAttachment({
        name: data.attachment.name,
        contentType: data.attachment.contentType || attachment.contentType,
        contentBytes: data.attachment.contentBytes,
      });
    } catch (uploadError) {
      setError(
        uploadError instanceof Error
          ? uploadError.message
          : "Failed to save Excel attachment.",
      );
      setQuestionListAttachment(null);
      event.target.value = "";
    } finally {
      setUploadingQuestionList(false);
    }
  }

  async function handleBd15FileUpload(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    if (files.length === 0 || !selected) {
      return;
    }
    if (!requireActiveSchedule("attach report files")) {
      event.target.value = "";
      return;
    }

    setUploadingBd15Files(true);
    setError(null);
    try {
      let nextAttachments = bd15Attachments;
      for (const file of files) {
        const lower = file.name.toLowerCase();
        const isPdf =
          lower.endsWith(".pdf") || file.type === "application/pdf";
        const isExcel =
          lower.endsWith(".xlsx") ||
          lower.endsWith(".xls") ||
          lower.endsWith(".csv") ||
          file.type.includes("spreadsheet") ||
          file.type.includes("excel") ||
          file.type === "text/csv";
        if (!isPdf && !isExcel) {
          throw new Error(
            `Only Excel or PDF files are allowed (${file.name}).`,
          );
        }

        const existing = nextAttachments.find(
          (item) => item.name.toLowerCase() === file.name.toLowerCase(),
        );
        let replaceSameName = false;
        if (existing) {
          const replace = window.confirm(
            `"${file.name}" is already uploaded. Replace the previous file?`,
          );
          if (!replace) {
            continue;
          }
          replaceSameName = true;
        }

        const buffer = await file.arrayBuffer();
        const attachment = {
          name: file.name,
          contentType: mimeTypeForBd15File(file),
          contentBytes: arrayBufferToBase64(buffer),
        };
        const response = await portalFetch("/api/bd15-email-attachments", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            clientId: selected.resource.id,
            monthKey: currentMonthKey(scheduleNow),
            replaceSameName,
            ...attachment,
          }),
        });
        const data = (await response.json()) as {
          attachments?: Bd15ApiAttachment[];
          error?: string;
        };
        if (!response.ok || !Array.isArray(data.attachments)) {
          throw new Error(data.error || "Failed to save PDF attachment.");
        }
        nextAttachments = mapBd15ApiAttachments(data.attachments);
      }
      setBd15Attachments(nextAttachments);
      event.target.value = "";
    } catch (uploadError) {
      setError(
        uploadError instanceof Error
          ? uploadError.message
          : "Failed to save PDF attachments.",
      );
      event.target.value = "";
    } finally {
      setUploadingBd15Files(false);
    }
  }

  async function handleRemoveBd15Attachment(attachmentId: string) {
    if (!selected) {
      return;
    }
    if (!requireActiveSchedule("remove PDF attachments")) {
      return;
    }
    setUploadingBd15Files(true);
    setError(null);
    try {
      const response = await portalFetch(
        `/api/bd15-email-attachments?clientId=${encodeURIComponent(selected.resource.id)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}&id=${encodeURIComponent(attachmentId)}`,
        { method: "DELETE" },
      );
      const data = (await response.json()) as {
        attachments?: Bd15ApiAttachment[];
        error?: string;
      };
      if (!response.ok || !Array.isArray(data.attachments)) {
        throw new Error(data.error || "Failed to remove PDF attachment.");
      }
      setBd15Attachments(mapBd15ApiAttachments(data.attachments));
    } catch (removeError) {
      setError(
        removeError instanceof Error
          ? removeError.message
          : "Failed to remove PDF attachment.",
      );
    } finally {
      setUploadingBd15Files(false);
    }
  }

  async function handleMoveRequestItem(itemId: string, direction: -1 | 1) {
    if (!selected || !activeRequestScheduleId) {
      return;
    }
    if (!requireActiveSchedule("edit the request list")) {
      return;
    }
    const items = [...activeList.requestItems];
    const monthItems = requestItemsForMonth(items, scheduleMonthKey);
    const index = monthItems.findIndex((item) => item.id === itemId);
    const target = index + direction;
    if (index < 0 || target < 0 || target >= monthItems.length) {
      return;
    }
    const reordered = [...monthItems];
    const [item] = reordered.splice(index, 1);
    reordered.splice(target, 0, item!);
    const monthIds = new Set(monthItems.map((entry) => entry.id));
    let cursor = 0;
    const nextItems = items.map((entry) =>
      monthIds.has(entry.id) ? reordered[cursor++]! : entry,
    );
    await persistList({
      ...activeList,
      clientId: selected.resource.id,
      scheduleId: activeRequestScheduleId,
      requestItems: nextItems,
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
    if (!requireActiveSchedule("edit the request list")) {
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
    if (!requireActiveSchedule("edit the request list")) {
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
        scheduleNow,
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
    if (!isSelectedScheduleActive) {
      setError("This business day is Inactive. Fill the star to activate it.");
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

  async function handleCheckInSharePoint() {
    if (!selected || !showBd5Content) {
      return;
    }
    if (!isSelectedScheduleActive) {
      setError("BD5 is Inactive for this client. Fill the star to activate it.");
      return;
    }
    if (!selectedPortalLink) {
      setError("This client has no Portal Link to check in SharePoint.");
      return;
    }
    if (!canEdit) {
      setError("Admin access is required to update action items.");
      return;
    }

    setCheckingSharePoint(true);
    setError(null);
    setSharePointCheckMessage(null);

    try {
      let account =
        instance.getActiveAccount() ??
        instance.getAllAccounts()[0] ??
        msalAccount;
      if (!account) {
        const result = await instance.loginPopup({
          ...sharePointRequest,
          prompt: "select_account",
        });
        account = result.account;
        if (account) {
          instance.setActiveAccount(account);
        }
      }
      if (!account) {
        throw new Error(
          "Sign in with a Fire Fly Microsoft account to check SharePoint.",
        );
      }

      const periodDate = scheduleNow;
      const folder = await listSharePointMonthFolderFiles(
        instance,
        account,
        selectedPortalLink,
        periodDate,
      );

      const bankFileFound = sharePointFileNameMatches(
        folder.fileNames,
        "Bank Statement",
      );
      const creditFileFound = sharePointFileNameMatches(
        folder.fileNames,
        "Credit Card Statement",
      );

      const monthKey = currentMonthKey(periodDate);
      const candidates = bd5ActionTableItems.filter(
        (item) => item.monthKey === monthKey || !item.monthKey,
      );

      let store = requestLists;
      let bankUpdated = 0;
      let creditUpdated = 0;
      const touchedLists = new Map<string, ClientRequestList>();

      for (const item of candidates) {
        const kind = classifySharePointActionItem(item.label);
        if (!kind) {
          continue;
        }
        const nextStatus =
          kind === "bank"
            ? bankFileFound
              ? "completed"
              : "late"
            : creditFileFound
              ? "completed"
              : "late";

        const owner =
          findListContainingActionItem(
            store,
            selected.resource.id,
            item.id,
          ) ??
          getClientRequestList(
            store,
            selected.resource.id,
            activeRequestScheduleId,
          );
        const updated = setActionItemStatus(owner, item.id, nextStatus);
        store = upsertClientRequestList(store, updated, scheduleNow);
        touchedLists.set(`${updated.clientId}::${updated.scheduleId}`, updated);
        if (kind === "bank") {
          bankUpdated += 1;
        } else {
          creditUpdated += 1;
        }
      }

      if (touchedLists.size === 0) {
        setSharePointCheckMessage(
          `Checked ${folder.folderPath}: Bank Statement ${bankFileFound ? "found" : "not found"}; Credit Card Statement ${creditFileFound ? "found" : "not found"}. No matching Bank/Credit Card Statement action items to update.`,
        );
        return;
      }

      setSavingRequestList(true);
      let nextStore = store;
      for (const list of touchedLists.values()) {
        const response = await portalFetch("/api/client-request-lists", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            ...list,
            monthKey: currentMonthKey(scheduleNow),
          }),
        });
        const data = await response.json();
        if (!response.ok) {
          throw new Error(
            (data as { error?: string }).error ||
              "Failed to save action item updates.",
          );
        }
        nextStore = normalizeClientRequestListsStore(data);
      }
      setRequestLists(nextStore);

      setSharePointCheckMessage(
        `Checked ${folder.folderPath}: Bank Statement ${bankFileFound ? "found → Completed" : "not found → Late"} (${bankUpdated}); Credit Card Statement ${creditFileFound ? "found → Completed" : "not found → Late"} (${creditUpdated}).`,
      );
    } catch (checkError) {
      setError(
        checkError instanceof Error
          ? checkError.message
          : "Failed to check SharePoint folder.",
      );
    } finally {
      setCheckingSharePoint(false);
      setSavingRequestList(false);
    }
  }

  async function handleReopenActionItem(itemId: string) {
    if (!selected) {
      return;
    }
    if (!requireActiveSchedule("update action items")) {
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

  async function markBd15EmailAsSent(sentAt: string = new Date().toISOString()) {
    if (!selected) {
      return;
    }
    try {
      const response = await portalFetch("/api/bd15-email-status", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "mark-sent",
          clientId: selected.resource.id,
          monthKey: currentMonthKey(scheduleNow),
          sentAt,
        }),
      });
      const data = (await response.json()) as {
        status?: {
          emailSent?: boolean;
          emailSentAt?: string;
          comments?: Array<{
            id?: string;
            text?: string;
            createdAt?: string;
          }>;
        };
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to save BD15 email status.");
      }
      const status = data.status;
      setBd15EmailStatus({
        emailSent: Boolean(status?.emailSent),
        emailSentAt:
          typeof status?.emailSentAt === "string" ? status.emailSentAt : sentAt,
        comments: Array.isArray(status?.comments)
          ? status.comments
              .filter(
                (item) =>
                  typeof item.id === "string" &&
                  item.id &&
                  typeof item.text === "string" &&
                  item.text.trim(),
              )
              .map((item) => ({
                id: item.id as string,
                text: (item.text as string).trim(),
                createdAt:
                  typeof item.createdAt === "string" ? item.createdAt : "",
              }))
          : bd15EmailStatus.comments,
      });
    } catch {
      // Email already went out; status update failure should not block the UI.
    }
  }

  async function handleAddBd18Comment() {
    const text = bd18CommentDraft.trim();
    if (!text || !selected) {
      return;
    }
    if (!requireActiveSchedule("add a status comment")) {
      return;
    }
    setSavingBd18Comment(true);
    setError(null);
    try {
      const response = await portalFetch("/api/bd15-email-status", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "add-comment",
          clientId: selected.resource.id,
          monthKey: currentMonthKey(scheduleNow),
          text,
        }),
      });
      const data = (await response.json()) as {
        status?: {
          emailSent?: boolean;
          emailSentAt?: string;
          comments?: Array<{
            id?: string;
            text?: string;
            createdAt?: string;
          }>;
        };
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to save comment.");
      }
      const status = data.status;
      setBd15EmailStatus({
        emailSent: Boolean(status?.emailSent),
        emailSentAt:
          typeof status?.emailSentAt === "string" ? status.emailSentAt : "",
        comments: Array.isArray(status?.comments)
          ? status.comments
              .filter(
                (item) =>
                  typeof item.id === "string" &&
                  item.id &&
                  typeof item.text === "string" &&
                  item.text.trim(),
              )
              .map((item) => ({
                id: item.id as string,
                text: (item.text as string).trim(),
                createdAt:
                  typeof item.createdAt === "string" ? item.createdAt : "",
              }))
          : [],
      });
      setBd18CommentDraft("");
    } catch (commentError) {
      setError(
        commentError instanceof Error
          ? commentError.message
          : "Failed to save comment.",
      );
    } finally {
      setSavingBd18Comment(false);
    }
  }

  async function handleSendEmailNow() {
    if (reminderEmailDisabled) {
      setEmailSendError(
        reminderEmailDisabledTitle || "No late items found for email.",
      );
      setEmailSendMessage(null);
      return;
    }
    if (!isSelectedScheduleActive) {
      setEmailSendError(
        "This business day is Inactive. Fill the star to send email.",
      );
      setEmailSendMessage(null);
      return;
    }
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
        attachments: showBd15Email
          ? bd15Attachments.map(({ name, contentType, contentBytes }) => ({
              name,
              contentType,
              contentBytes,
            }))
          : showBd8Content && questionListAttachment
            ? [questionListAttachment]
            : undefined,
      });
      if (showBd15Email) {
        await markBd15EmailAsSent();
      }
      setEmailSendMessage(
        showBd15Email && bd15Attachments.length > 0
          ? `Email sent with ${bd15Attachments.length} attachment${bd15Attachments.length === 1 ? "" : "s"}.`
          : showBd8Content && questionListAttachment
            ? `Email sent with attachment (${questionListAttachment.name}).`
            : "Email sent.",
      );
    } catch (sendError) {
      setEmailSendError(formatMailSendError(sendError));
      setEmailSendMessage(null);
    } finally {
      setSendingEmail(false);
    }
  }

  async function handleScheduleEmail() {
    if (reminderEmailDisabled) {
      setEmailSendError(
        reminderEmailDisabledTitle || "No late items found for email.",
      );
      setEmailSendMessage(null);
      return;
    }
    if (!isSelectedScheduleActive) {
      setEmailSendError(
        "This business day is Inactive. Fill the star to schedule email.",
      );
      setEmailSendMessage(null);
      return;
    }
    if (!emailTo.trim() || !emailSubject.trim()) {
      setEmailSendError("To and Subject are required.");
      setEmailSendMessage(null);
      return;
    }
    if (clientTimeZoneError) {
      setEmailSendError(clientTimeZoneError);
      setEmailSendMessage(null);
      return;
    }

    const sendAt = parseDateTimeLocalInTimeZone(
      effectiveEmailSendAt,
      USER_TIME_ZONE,
    );
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
        attachments: showBd15Email
          ? bd15Attachments.map(({ name, contentType, contentBytes }) => ({
              name,
              contentType,
              contentBytes,
            }))
          : showBd8Content && questionListAttachment
            ? [questionListAttachment]
            : undefined,
        sendAt,
      });
      const whenLabel = new Intl.DateTimeFormat("en-GB", {
        timeZone: USER_TIME_ZONE,
        weekday: "long",
        day: "numeric",
        month: "short",
        year: "numeric",
        hour: "numeric",
        minute: "2-digit",
        hour12: true,
      }).format(sendAt);
      if (showBd15Email) {
        await markBd15EmailAsSent(sendAt.toISOString());
      }
      setEmailSendMessage(
        showBd15Email && bd15Attachments.length > 0
          ? `Email scheduled for ${whenLabel} (Pakistan) with ${bd15Attachments.length} attachment${bd15Attachments.length === 1 ? "" : "s"}.`
          : showBd8Content && questionListAttachment
            ? `Email scheduled for ${whenLabel} (Pakistan) with attachment (${questionListAttachment.name}).`
            : `Email scheduled for ${whenLabel} (Pakistan).`,
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
          <p
            className={
              demoMonthLabel
                ? "client-management-demo-banner"
                : "client-management-demo-banner client-management-demo-banner--empty"
            }
            role="status"
          >
            {demoMonthLabel
              ? `Demo month: ${demoMonthLabel} — BD dates follow this month (set in Lookup → Repeated Tasks).`
              : "\u00a0"}
          </p>
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
                                  {left.key === "bd1Template" ? (
                                    selected.profile ? (
                                      <select
                                        className="field-input client-info-select"
                                        value={selectedBd1Template}
                                        disabled={!canEdit || savingBd1Template}
                                        onChange={(event) =>
                                          void handleBd1TemplateChange(
                                            normalizeBd1Template(
                                              event.target.value,
                                            ),
                                          )
                                        }
                                      >
                                        {BD1_TEMPLATE_OPTIONS.map((option) => (
                                          <option
                                            key={option.id}
                                            value={option.id}
                                          >
                                            {option.label}
                                          </option>
                                        ))}
                                      </select>
                                    ) : (
                                      "—"
                                    )
                                  ) : (
                                    left.value
                                  )}
                                </td>
                                <th scope="row" className="client-info-field">
                                  {right?.label ?? ""}
                                </th>
                                <td className="client-info-value">
                                  {right?.key === "bd1Template" ? (
                                    selected.profile ? (
                                      <select
                                        className="field-input client-info-select"
                                        value={selectedBd1Template}
                                        disabled={!canEdit || savingBd1Template}
                                        onChange={(event) =>
                                          void handleBd1TemplateChange(
                                            normalizeBd1Template(
                                              event.target.value,
                                            ),
                                          )
                                        }
                                      >
                                        {BD1_TEMPLATE_OPTIONS.map((option) => (
                                          <option
                                            key={option.id}
                                            value={option.id}
                                          >
                                            {option.label}
                                          </option>
                                        ))}
                                      </select>
                                    ) : (
                                      "—"
                                    )
                                  ) : (
                                    right?.value ?? ""
                                  )}
                                </td>
                              </tr>
                            ),
                          )}
                        </tbody>
                      </table>
                    </div>
                  ) : null}
                </div>

                <div className="client-schedule-layout">
                  <aside
                    className="client-schedule-sidebar"
                    aria-label="Client schedule"
                  >
                    <div className="schedule-pane-header">
                      <p className="schedule-pane-title">Active</p>
                      <p className="schedule-pane-hint">
                        Only these days send email and run actions.
                      </p>
                    </div>

                    <div className="schedule-pane-list">
                      {activeScheduleItems.length === 0 ? (
                        <p className="schedule-section-empty">
                          No active days. Expand Inactive below and click ☆ to
                          activate.
                        </p>
                      ) : (
                        activeScheduleItems.map((item) => {
                          const businessDate = scheduleDateById.get(item.id);
                          const isSelected = selectedSchedule.id === item.id;
                          return (
                            <div
                              key={item.id}
                              className={
                                isSelected
                                  ? "schedule-row schedule-row--on schedule-row--selected"
                                  : "schedule-row schedule-row--on"
                              }
                            >
                              <button
                                type="button"
                                className="schedule-star schedule-star--filled"
                                aria-label={`Deactivate ${item.label}`}
                                title="Move to Inactive"
                                disabled={!canEdit || savingScheduleSettings}
                                onClick={(event) =>
                                  void handleToggleScheduleStar(item.id, event)
                                }
                              >
                                ★
                              </button>
                              <button
                                type="button"
                                className="schedule-day-link"
                                onClick={() => setSelectedScheduleId(item.id)}
                              >
                                <span className="client-schedule-link-label">
                                  {item.label}
                                </span>
                                {businessDate ? (
                                  <span className="client-schedule-link-date">
                                    {businessDate}
                                  </span>
                                ) : null}
                              </button>
                            </div>
                          );
                        })
                      )}
                    </div>

                    <div className="schedule-inactive-block">
                      {inactiveScheduleItems.length === 0 ? (
                        <p className="schedule-inactive-toggle-label">
                          Inactive · none
                        </p>
                      ) : (
                        <button
                          type="button"
                          className="schedule-inactive-toggle"
                          aria-expanded={inactiveExpanded}
                          onClick={() =>
                            setInactiveExpanded((open) => !open)
                          }
                        >
                          <span>
                            Inactive · {inactiveScheduleItems.length}
                          </span>
                          <span aria-hidden="true">
                            {inactiveExpanded ? "▾" : "▸"}
                          </span>
                        </button>
                      )}

                      {inactiveExpanded && inactiveScheduleItems.length > 0 ? (
                        <div className="schedule-pane-list schedule-pane-list--inactive">
                          {inactiveScheduleItems.map((item) => {
                            const businessDate = scheduleDateById.get(item.id);
                            const isSelected =
                              selectedSchedule.id === item.id;
                            return (
                              <div
                                key={item.id}
                                className={
                                  isSelected
                                    ? "schedule-row schedule-row--off schedule-row--selected"
                                    : "schedule-row schedule-row--off"
                                }
                              >
                                <button
                                  type="button"
                                  className="schedule-star"
                                  aria-label={`Activate ${item.label}`}
                                  title="Move to Active"
                                  disabled={
                                    !canEdit || savingScheduleSettings
                                  }
                                  onClick={(event) =>
                                    void handleToggleScheduleStar(
                                      item.id,
                                      event,
                                    )
                                  }
                                >
                                  ☆
                                </button>
                                <button
                                  type="button"
                                  className="schedule-day-link"
                                  onClick={() =>
                                    setSelectedScheduleId(item.id)
                                  }
                                >
                                  <span className="client-schedule-link-label">
                                    {item.label}
                                  </span>
                                  {businessDate ? (
                                    <span className="client-schedule-link-date">
                                      {businessDate}
                                    </span>
                                  ) : null}
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      ) : null}
                    </div>
                  </aside>

                  <div className="client-schedule-panel">
                    {!isSelectedScheduleActive ? (
                      <p className="inline-notice inline-notice--error schedule-inactive-banner">
                        <span className="inline-notice__text">
                          This business day is Inactive. Fill the star in the
                          sidebar to activate it before sending email or updating
                          actions.
                        </span>
                      </p>
                    ) : null}
                    {(() => {
                      const workflowTabsClassName = showRequestTypeChooser
                        ? "lookup-section-tabs bd1-workflow-tabs request-workflow-tabs"
                        : "lookup-section-tabs request-workflow-tabs";

                      const requestWorkflow = (
                        <>
                          {showLastBdSchedule ? (
                            <div
                              className="lookup-section-tabs request-workflow-tabs last-bd-audience-tabs"
                              role="tablist"
                              aria-label="Email audience"
                            >
                              <button
                                type="button"
                                role="tab"
                                aria-selected={lastBdAudience === "client"}
                                className={
                                  lastBdAudience === "client"
                                    ? "lookup-section-tab active"
                                    : "lookup-section-tab"
                                }
                                onClick={() => setLastBdAudience("client")}
                              >
                                Client
                              </button>
                              <button
                                type="button"
                                role="tab"
                                aria-selected={lastBdAudience === "team"}
                                className={
                                  lastBdAudience === "team"
                                    ? "lookup-section-tab active"
                                    : "lookup-section-tab"
                                }
                                onClick={() => setLastBdAudience("team")}
                              >
                                Team
                              </button>
                            </div>
                          ) : null}
                          {!showBd5Content && !showEndOfMonthEmailOnly ? (
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
                          {showBd15Email
                            ? "Attachments"
                            : showBd8Content
                              ? "Question List Template"
                              : "Request List Template"}
                        </button>
                      ) : null}
                      {showBd9LikeContent ||
                      (!showBd1LikeRequestList && !showBd5Content) ? (
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
                        {showBd9LikeContent
                          ? "Question List Status"
                          : showBd7Content
                            ? "Late Action Items"
                            : "Action Items"}
                      </button>
                      ) : null}
                      {showBd1LikeWorkflow ? (
                        <button
                          type="button"
                          role="tab"
                          aria-selected={workflowTab === "email"}
                          aria-disabled={reminderEmailTabDisabled || undefined}
                          disabled={reminderEmailTabDisabled}
                          title={
                            reminderEmailTabDisabled
                              ? reminderEmailDisabledTitle
                              : undefined
                          }
                          className={
                            workflowTab === "email"
                              ? "lookup-section-tab active"
                              : "lookup-section-tab"
                          }
                          onClick={() => {
                            if (reminderEmailTabDisabled) {
                              return;
                            }
                            setWorkflowTab("email");
                          }}
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
                          ) : null}
                          {showBd7Content ? (
                            <p className="bd1-workflow-kicker">
                              Late Action Items shows only items marked Late on
                              BD5 for the selected request type.
                            </p>
                          ) : null}
                    {showRequestListTab && workflowTab === "request" ? (
                      <div className="request-list-panel">
                        {showBd15Email ? (
                          <>
                            {(() => {
                              const sentAtLabel = formatActionResolvedAt(
                                bd15EmailStatus.emailSentAt,
                              );
                              return (
                                <p
                                  className={
                                    bd15EmailStatus.emailSent
                                      ? "bd18-email-status bd18-email-status--sent"
                                      : "bd18-email-status bd18-email-status--not-sent"
                                  }
                                >
                                  {bd15EmailStatus.emailSent
                                    ? `Email sent${sentAtLabel ? ` on ${sentAtLabel}` : ""}`
                                    : "Email not sent"}
                                </p>
                              );
                            })()}
                            <p className="lookup-hint">
                              Upload the required PDF files for this month. They
                              are kept when you switch tabs and cleared
                              automatically next month.
                            </p>
                            <ul className="bd15-required-reports">
                              {BD15_REQUIRED_REPORTS.map((report) => {
                                const found = bd15Attachments.some((file) =>
                                  bd15FileMatchesReport(file.name, report),
                                );
                                return (
                                  <li
                                    key={report}
                                    className={
                                      found
                                        ? "bd15-required-report bd15-required-report--found"
                                        : "bd15-required-report bd15-required-report--missing"
                                    }
                                  >
                                    {found ? "✓" : "○"} {report}
                                  </li>
                                );
                              })}
                            </ul>
                            {bd15MissingReports.length > 0 ? (
                              <p className="form-message error">
                                These files are missing:{" "}
                                {bd15MissingReports.join(", ")}
                              </p>
                            ) : null}
                            <div className="question-list-upload">
                              <label className="field field--wide">
                                <span className="sr-only">
                                  Books closed report files
                                </span>
                                <input
                                  ref={bd15FileRef}
                                  className="field-input"
                                  type="file"
                                  accept=".xlsx,.xls,.csv,.pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv,application/pdf"
                                  multiple
                                  disabled={
                                    uploadingBd15Files ||
                                    !isSelectedScheduleActive ||
                                    bd15EmailAlreadySent
                                  }
                                  onChange={(event) =>
                                    void handleBd15FileUpload(event)
                                  }
                                />
                              </label>
                              <p className="lookup-hint">
                                {uploadingBd15Files
                                  ? "Saving…"
                                  : "Accepts multiple Excel (.xlsx, .xls, .csv) or PDF files. File names should include the report name."}
                              </p>
                            </div>
                            {bd15Attachments.length === 0 ? (
                              <p className="lookup-hint">
                                No attachments yet.
                              </p>
                            ) : (
                              <div className="table-wrap bd5-action-table-wrap">
                                <table className="projects-table bd5-action-table">
                                  <thead>
                                    <tr>
                                      <th scope="col">List</th>
                                      <th scope="col">Date Time</th>
                                      <th scope="col">Action</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {bd15Attachments.map((file) => {
                                      const uploadedAt = formatActionResolvedAt(
                                        file.savedAt,
                                      );
                                      return (
                                        <tr
                                          key={file.id}
                                          className="bd5-action-row"
                                        >
                                          <td className="bd5-action-list-cell">
                                            {file.name}
                                          </td>
                                          <td className="bd5-action-datetime-cell">
                                            {uploadedAt || "—"}
                                          </td>
                                          <td>
                                            <button
                                              type="button"
                                              className="button secondary"
                                              disabled={
                                                uploadingBd15Files ||
                                                !isSelectedScheduleActive ||
                                                bd15EmailAlreadySent
                                              }
                                              onClick={() =>
                                                void handleRemoveBd15Attachment(
                                                  file.id,
                                                )
                                              }
                                            >
                                              Remove
                                            </button>
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            )}

                            {showBd18Content ? (
                              <div className="bd18-comments">
                                <h5 className="bd18-comments-title">
                                  Comments
                                </h5>
                                <p className="lookup-hint">
                                  {bd15EmailStatus.emailSent
                                    ? "Notes about this BD15 email (saved with date and time)."
                                    : "Add a note explaining why BD15 email was not sent (saved with date and time)."}
                                </p>
                                <div className="request-list-add">
                                  <input
                                    className="field-input"
                                    value={bd18CommentDraft}
                                    placeholder="Why was this not sent on BD15?"
                                    disabled={
                                      savingBd18Comment ||
                                      !isSelectedScheduleActive
                                    }
                                    onChange={(event) =>
                                      setBd18CommentDraft(event.target.value)
                                    }
                                    onKeyDown={(event) => {
                                      if (event.key === "Enter") {
                                        event.preventDefault();
                                        void handleAddBd18Comment();
                                      }
                                    }}
                                  />
                                  <button
                                    type="button"
                                    className="button secondary-solid"
                                    disabled={
                                      savingBd18Comment ||
                                      !isSelectedScheduleActive ||
                                      !bd18CommentDraft.trim()
                                    }
                                    onClick={() => void handleAddBd18Comment()}
                                  >
                                    {savingBd18Comment
                                      ? "Saving…"
                                      : "Add comment"}
                                  </button>
                                </div>
                                {bd15EmailStatus.comments.length === 0 ? (
                                  <p className="lookup-hint">
                                    No comments yet.
                                  </p>
                                ) : (
                                  <div className="table-wrap bd5-action-table-wrap">
                                    <table className="projects-table bd5-action-table">
                                      <thead>
                                        <tr>
                                          <th scope="col">Comment</th>
                                          <th scope="col">Date Time</th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {bd15EmailStatus.comments.map(
                                          (comment) => (
                                            <tr
                                              key={comment.id}
                                              className="bd5-action-row"
                                            >
                                              <td className="bd5-action-list-cell">
                                                {comment.text}
                                              </td>
                                              <td className="bd5-action-datetime-cell">
                                                {formatActionResolvedAt(
                                                  comment.createdAt,
                                                ) || "—"}
                                              </td>
                                            </tr>
                                          ),
                                        )}
                                      </tbody>
                                    </table>
                                  </div>
                                )}
                              </div>
                            ) : null}
                          </>
                        ) : (
                          <>
                        <p className="lookup-hint">
                          {showBd8Content
                            ? "Add questions as text below. Excel is saved as the email attachment for this month (kept when you switch tabs; cleared automatically next month)."
                            : "Add items here. The Email tab uses these bullets directly for this request type."}
                        </p>
                        {showBd8Content ? (
                          <div className="question-list-upload">
                            <label className="field field--wide">
                              <span className="sr-only">
                                Question list Excel file
                              </span>
                              <input
                                ref={questionListFileRef}
                                className="field-input"
                                type="file"
                                accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv"
                                disabled={
                                  uploadingQuestionList ||
                                  savingRequestList ||
                                  !isSelectedScheduleActive
                                }
                                onChange={(event) =>
                                  void handleQuestionListExcelUpload(event)
                                }
                              />
                            </label>
                            <p className="lookup-hint">
                              {uploadingQuestionList
                                ? "Loading…"
                                : questionListAttachment
                                  ? `Email attachment: ${questionListAttachment.name}`
                                  : "Accepts .xlsx, .xls, or .csv. Saved until replaced or the month changes."}
                            </p>
                          </div>
                        ) : null}
                        <div className="request-list-add">
                          <input
                            className="field-input"
                            value={draftRequestLabel}
                            placeholder={
                              showBd8Content
                                ? "Add a question…"
                                : "Add bullet point…"
                            }
                            disabled={
                              savingRequestList || !isSelectedScheduleActive
                            }
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
                              savingRequestList ||
                              !isSelectedScheduleActive ||
                              !draftRequestLabel.trim()
                            }
                            onClick={() => void handleAddRequestItem()}
                          >
                            Add
                          </button>
                        </div>

                        {visibleRequestItems.length === 0 ? (
                          <p className="lookup-hint">
                            {showBd8Content
                              ? "No questions loaded yet."
                              : "No request bullets yet."}
                          </p>
                        ) : (
                          <ul className="request-list-items">
                            {visibleRequestItems.map((item, index) => {
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
                                              visibleRequestItems.length - 1 ||
                                            savingRequestList
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
                          </>
                        )}
                      </div>
                    ) : null}

                    {workflowTab === "action" ? (
                      <div className="request-list-panel">
                        {showReviewActionTable ? (
                          <>
                            {showBd5Content ? (
                              <>
                                <div className="bd5-action-toolbar">
                                  <p className="lookup-hint">
                                    Checks the same SharePoint month folder as
                                    the BD1 email portal link for files named
                                    Bank Statement / Credit Card Statement, then
                                    marks matching action rows Completed or Late.
                                  </p>
                                  <button
                                    type="button"
                                    className="button secondary-solid"
                                    disabled={
                                      !isSelectedScheduleActive ||
                                      checkingSharePoint ||
                                      savingRequestList ||
                                      !canEdit ||
                                      !selectedPortalLink
                                    }
                                    onClick={() =>
                                      void handleCheckInSharePoint()
                                    }
                                  >
                                    {checkingSharePoint
                                      ? "Checking SharePoint…"
                                      : "Check in SharePoint"}
                                  </button>
                                </div>
                                {sharePointCheckMessage ? (
                                  <p className="lookup-hint bd5-sharepoint-result">
                                    {sharePointCheckMessage}
                                  </p>
                                ) : null}
                              </>
                            ) : showBd9Content ? (
                              <p className="lookup-hint">
                                BD8 question list items. Mark each row Completed
                                or Late. Late items automatically appear on
                                BD10.
                              </p>
                            ) : null}
                            {bd5ActionTableItems.length === 0 ? (
                              <p className="lookup-hint">
                                {showBd10Content
                                  ? "No late BD9 questions found."
                                  : showBd9Content
                                    ? "No BD8 question action items yet."
                                    : "No action items for this type yet."}
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
                                            ) : showBd10Content &&
                                              item.status === "late" ? (
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
                            {bd7LateItems.length === 0 ? (
                              <p className="lookup-hint">
                                No late action items for BD5 found.
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
                    !showReviewActionTable &&
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
                        {bd15EmailAlreadySent ? (
                          <p className="bd18-email-status bd18-email-status--sent">
                            Email sent on{" "}
                            {formatActionResolvedAt(
                              bd15EmailStatus.emailSentAt,
                            ) || "—"}
                          </p>
                        ) : null}
                        <p className="lookup-hint">
                          {showBd15Email ? (
                            <>
                              Subject and body are filled from Email Template 11.
                              All required report files from Attachments are
                              sent with the email
                              {bd15Attachments.length > 0
                                ? ` (${bd15Attachments.length} file${bd15Attachments.length === 1 ? "" : "s"})`
                                : ""}
                              {bd15EmailAlreadySent
                                ? ". Sending is disabled because this email was already sent."
                                : "."}
                            </>
                          ) : showBd10Content ? (
                            <>
                              Subject and body are filled from Email Template 7.
                              Outstanding items are only questions marked Late
                              on BD9.
                            </>
                          ) : showBd9Content ? (
                            <>
                              Subject and body are filled from Email Template 8.
                              [Number] uses open and late BD8 questions from
                              Question List Status.
                            </>
                          ) : showBd8Content ? (
                            <>
                              Subject and body are filled from Email Template 3.
                              [Number] and [Question List] come from questions
                              you add as text. The Excel file is email
                              attachment only
                              {questionListAttachment
                                ? ` (${questionListAttachment.name})`
                                : ""}.
                            </>
                          ) : showLast3BdContent ? (
                            <>
                              Subject and body are filled from Email Template 5.
                            </>
                          ) : showLast2BdContent ? (
                            <>
                              Subject and body are filled from Email Template 12.
                            </>
                          ) : showLast1BdClientContent ? (
                            <>
                              Client email for Last 1 BD — draft freely (no
                              numbered template yet). Switch to Team for
                              Template 18.
                            </>
                          ) : showLastBdTeamContent ? (
                            <>
                              Subject and body are filled from Email Template 18
                              (team end-of-month reminder). To is you; BCC is
                              everyone in User Management.
                            </>
                          ) : (
                            <>
                              Subject and body are filled from Email Template{" "}
                              {requestTypeId === "other-statements" ? "2" : "1"}
                              {showBd1TypeContent
                                ? ". The account list comes from the Request List Template."
                                : "."}
                            </>
                          )}
                        </p>
                        <div className="client-schedule-email-fields">
                          <label className="field">
                            <span className="field-label">To</span>
                            <input
                              className="field-input"
                              type="text"
                              inputMode="email"
                              value={emailTo}
                              disabled={bd15EmailAlreadySent}
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
                              disabled={bd15EmailAlreadySent}
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
                              disabled={bd15EmailAlreadySent}
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
                              disabled={bd15EmailAlreadySent}
                              onChange={(event) =>
                                setEmailDeadline(event.target.value)
                              }
                            />
                          </label>
                          <div className="field">
                            <span className="field-label">
                              Send at (client time zone)
                            </span>
                            <div className="field-input field-input--readonly">
                              {emailSendAtClientLabel || "—"}
                            </div>
                            {clientTimeZoneError ? (
                              <span className="form-message error">
                                {clientTimeZoneError}
                              </span>
                            ) : (
                              <span className="lookup-hint">
                                This BD date at 9:00 am in the client’s time
                                zone
                              </span>
                            )}
                          </div>
                          <label className="field">
                            <span className="field-label">
                              Scheduled send (Pakistan)
                            </span>
                            <input
                              className="field-input"
                              type="datetime-local"
                              value={effectiveEmailSendAt}
                              disabled={
                                bd15EmailAlreadySent ||
                                Boolean(clientTimeZoneError)
                              }
                              onChange={(event) =>
                                setEmailSendAt(event.target.value)
                              }
                            />
                            <span className="lookup-hint">
                              Same moment in Pakistan — matches Outlook drafts
                            </span>
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
                              reminderEmailDisabled ||
                              !isSelectedScheduleActive ||
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
                              reminderEmailDisabled ||
                              !isSelectedScheduleActive ||
                              sendingEmail ||
                              schedulingEmail ||
                              !emailTo.trim() ||
                              !emailSubject.trim() ||
                              !effectiveEmailSendAt.trim() ||
                              Boolean(clientTimeZoneError)
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

                      if (
                        !showRequestTypeChooser &&
                        !showBd8Content &&
                        !showBd9LikeContent &&
                        !showBd15Content &&
                        !showBd18Content
                      ) {
                        return requestWorkflow;
                      }

                      const requestBodyTitle = showBd15Email
                        ? "Books Closed."
                        : showBd10Content
                          ? "SECOND REMINDER: Uncategorized Transactions: Questions for the Client"
                          : showBd9Content
                            ? "FIRST REMINDER: Uncategorized Transactions: Questions for the Client"
                            : showBd8Content
                              ? "Uncategorized Transactions: Questions for the Client"
                              : activeRequestType.label;

                      return (
                        <div className="bd1-request-type-block">
                          <div
                            className="bd1-request-body"
                            role="region"
                            aria-labelledby="bd1-request-body-title"
                          >
                            <div className="bd1-request-body-header">
                              <div className="bd1-request-body-header-main">
                                <h4
                                  id="bd1-request-body-title"
                                  className="bd1-request-body-title"
                                >
                                  {requestBodyTitle}
                                </h4>
                                {showBd1TypeContent ? (
                                  <p className="bd1-request-body-hint">
                                    Set BD1 Template under Client Info. Template
                                    1 uses Bank and Credit Card; Template 2 uses
                                    Bank, Credit Card And Other Statements.
                                  </p>
                                ) : null}
                              </div>
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
