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
import {
  formatBusinessDayLabel,
  getScheduleBusinessDate,
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
  getClientActionItems,
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
import { readLocalAuthSession } from "@/lib/local-auth";
import { portalFetch } from "@/lib/portal-api";
import type { PortalUser } from "@/lib/portal-users";
import type { ResourceClientEntry } from "@/lib/resource-clients";
import { sendGraphMail } from "@/lib/send-graph-mail";

/** Bold auto-filled values (client name, period, etc.) inside subject/body. */
function renderAutofilledText(
  text: string,
  highlights: string[],
): ReactNode {
  if (!text) {
    return null;
  }
  const unique = [
    ...new Set(highlights.map((value) => value.trim()).filter(Boolean)),
  ].sort((a, b) => b.length - a.length);
  if (unique.length === 0) {
    return text;
  }
  const pattern = unique
    .map((value) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join("|");
  const splitter = new RegExp(`(${pattern})`, "g");
  const isHighlight = new Set(unique);
  return text.split(splitter).map((part, index) =>
    isHighlight.has(part) ? (
      <strong key={`${part}-${index}`} className="email-template-autofill">
        {part}
      </strong>
    ) : (
      <span key={`${index}-${part.slice(0, 12)}`}>{part}</span>
    ),
  );
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
  { key: "country", label: "Country" },
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

/** Schedule list ids that belong to a BD1/BD5 request type. */
function requestTypeScheduleIds(subMenuId: Bd1SubMenuId): string[] {
  return subMenuId === "other-statements"
    ? ["bd1-other-statements", "bd5-other-statements"]
    : ["bd1", "bd5"];
}

function getActionItemsForRequestType(
  store: ClientRequestListsStore,
  clientId: string,
  subMenuId: Bd1SubMenuId,
): ActionWorkflowItem[] {
  const allowed = new Set(requestTypeScheduleIds(subMenuId));
  const byKey = new Map<string, ActionWorkflowItem>();
  for (const list of store.lists) {
    if (list.clientId !== clientId || !allowed.has(list.scheduleId)) {
      continue;
    }
    for (const item of list.actionItems) {
      const key = `${item.sourceRequestId || item.id}::${item.monthKey}`;
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
    case "country":
      return profile?.country?.trim() || "—";
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
  const { email } = useAccess();
  const msalAccount = accounts[0] ?? null;
  const msalName = msalAccount?.name ?? "";

  const [resources, setResources] = useState<ResourceClientEntry[]>([]);
  const [profiles, setProfiles] = useState<ClientProfileEntry[]>([]);
  const [portalName, setPortalName] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedScheduleId, setSelectedScheduleId] = useState<string>(
    CLIENT_SCHEDULE_ITEMS[0].id,
  );
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
  const [emailFrom, setEmailFrom] = useState("");
  const [emailCc, setEmailCc] = useState("");
  const [emailBcc, setEmailBcc] = useState("");
  const [emailDeadline, setEmailDeadline] = useState("");
  const [emailSubject, setEmailSubject] = useState("");
  const [emailBody, setEmailBody] = useState("");
  const [emailAutofillHighlights, setEmailAutofillHighlights] = useState<
    string[]
  >([]);
  const [clientInfoOpen, setClientInfoOpen] = useState(false);
  const [sendingEmail, setSendingEmail] = useState(false);
  const [emailSendMessage, setEmailSendMessage] = useState<string | null>(null);
  const [emailSendError, setEmailSendError] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const showBd1TypeContent = selectedScheduleId === "bd1";
  const showBd5Content = selectedScheduleId === "bd5";
  const showBd7Content = selectedScheduleId === "bd7";
  const showBd1LikeWorkflow = showBd1TypeContent || showBd7Content;
  const showRequestTypeChooser =
    selectedScheduleId === "bd1" ||
    selectedScheduleId === "bd5" ||
    selectedScheduleId === "bd7";
  const showRequestListTab = showBd1TypeContent;
  const activeRequestType =
    BD1_SUB_MENUS.find((item) => item.id === requestTypeId) ??
    BD1_SUB_MENUS[0];
  const activeRequestScheduleId = showRequestTypeChooser
    ? requestTypeScheduleId(selectedScheduleId, requestTypeId)
    : selectedScheduleId;

  useEffect(() => {
    setEditingRequestItemId(null);
    setEditRequestLabel("");
  }, [selectedId, selectedScheduleId, workflowTab, requestTypeId]);

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
    if (showRequestTypeChooser) {
      return;
    }
    setRequestTypeId("bank-credit-card");
  }, [showRequestTypeChooser, selectedScheduleId]);

  useEffect(() => {
    if (showBd5Content && workflowTab === "completed") {
      setWorkflowTab("action");
    }
  }, [showBd5Content, workflowTab]);

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
      ] = await Promise.all([
        fetch("/api/resource-clients", { cache: "no-store" }),
        fetch("/api/client-profiles", { cache: "no-store" }),
        fetch("/api/portal-users", { cache: "no-store" }),
        fetch("/api/client-request-lists", { cache: "no-store" }),
        fetch("/api/holidays", { cache: "no-store" }),
        fetch("/api/email-templates", { cache: "no-store" }),
        portalFetch("/api/email-signature", { cache: "no-store" }),
      ]);
      const resourcesData = (await resourcesRes.json()) as ResourcesResponse;
      const profilesData = (await profilesRes.json()) as ProfilesResponse;
      const usersData = (await usersRes.json()) as UsersResponse;
      const requestListsData = await requestListsRes.json();
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

  const defaultEmailDeadline = useMemo(() => {
    if (
      !selectedClientId ||
      (selectedScheduleId !== "bd1" && selectedScheduleId !== "bd7")
    ) {
      return "";
    }
    const bd4Date = getScheduleBusinessDate(
      "bd4",
      holidays,
      selected?.profile?.country,
      selected?.profile?.state ?? "",
    );
    if (!bd4Date) {
      return "";
    }
    const year = bd4Date.getFullYear();
    const month = String(bd4Date.getMonth() + 1).padStart(2, "0");
    const day = String(bd4Date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }, [
    selectedClientId,
    selectedScheduleId,
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
      (selectedScheduleId !== "bd1" && selectedScheduleId !== "bd7")
    ) {
      return;
    }
    const template = findBd1EmailTemplate(emailTemplates, requestTypeId);
    const clientName =
      selectedOperatingName.replace(/[.:]+$/g, "").trim() ||
      selectedOperatingName.trim() ||
      "[Client Name]";
    const contactName = selectedPointOfContact || "[Contact First Name]";
    const portalLink = selectedPortalLink || "[Portal Link]";
    const period = formatCurrentPeriod();
    const dueDate = emailDeadlineLabel;
    const typeActionItems = getActionItemsForRequestType(
      requestLists,
      selectedClientId,
      requestTypeId,
    );
    const emailListItems =
      selectedScheduleId === "bd7"
        ? typeActionItems.filter((item) => item.status === "late")
        : getClientRequestList(
            requestLists,
            selectedClientId,
            requestTypeScheduleId("bd1", requestTypeId),
          ).actionItems.filter((item) => item.status === "action");
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
        text
          .replace(/\[Client Name\]/gi, clientName)
          .replace(/\[Contact First Name\]/gi, contactName)
          .replace(/\[Due Date\]/gi, dueDate)
          .replace(/\[Deadline\]/gi, dueDate)
          .replace(/\[Portal Link\]/gi, portalLink)
          .replace(
            /(?:•\s*)?\[List of bank and credit card accounts, with last 4 digits\]/gi,
            actionListText,
          )
          .replace(
            /(?:•\s*)?\[List of loan, line of credit, lease, investment or other accounts\]/gi,
            actionListText,
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
      portalLink,
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
    setEmailAutofillHighlights(highlights);
  }, [
    selectedClientId,
    selectedScheduleId,
    requestTypeId,
    selectedOperatingName,
    selectedPointOfContact,
    selectedPortalLink,
    emailDeadlineLabel,
    requestLists,
    emailTemplates,
  ]);

  useEffect(() => {
    if (
      !selectedClientId ||
      (selectedScheduleId !== "bd1" && selectedScheduleId !== "bd7")
    ) {
      return;
    }
    // Leave To / From empty on load for testing.
    setEmailTo("");
    setEmailFrom("");
    setEmailCc("");
    setEmailBcc("");
    setEmailDeadline("");
    setEmailSendError(null);
    setEmailSendMessage(null);
  }, [selectedClientId, selectedScheduleId, requestTypeId]);

  const selectedSchedule =
    CLIENT_SCHEDULE_ITEMS.find((item) => item.id === selectedScheduleId) ??
    CLIENT_SCHEDULE_ITEMS[0];

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
      );
      if (date) {
        map.set(item.id, formatBusinessDayLabel(date));
      }
    }
    return map;
  }, [selected, holidays]);

  const selectedScheduleDate =
    scheduleDateById.get(selectedSchedule.id) ?? null;

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

  const sharedActionItems = selected
    ? getClientActionItems(requestLists, selected.resource.id)
    : [];
  const scopedActionItems =
    showRequestTypeChooser && selected
      ? getActionItemsForRequestType(
          requestLists,
          selected.resource.id,
          requestTypeId,
        )
      : sharedActionItems;
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
      let account =
        instance.getActiveAccount() ??
        instance.getAllAccounts()[0] ??
        msalAccount;

      // ID/password portal login does not create a Graph token. If this is also
      // a Microsoft account, connect it now (login hint = portal email).
      if (!account) {
        const hint =
          email?.trim() ||
          readLocalAuthSession()?.email ||
          emailFrom.trim() ||
          "";
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
          "Microsoft sign-in was cancelled. Sign in with the Microsoft account that owns this mailbox, then try Send Now again.",
        );
      }

      setEmailSendMessage(null);
      await sendGraphMail(instance, account, {
        to: emailTo,
        cc: emailCc,
        bcc: emailBcc,
        subject: emailSubject,
        bodyText: emailBody,
        highlights: emailAutofillHighlights,
        signature: emailSignature,
      });
      setEmailSendMessage("Email sent.");
    } catch (sendError) {
      const raw =
        sendError instanceof Error
          ? sendError.message
          : "Failed to send email.";
      const lower = raw.toLowerCase();
      if (
        lower.includes("aadsts50020") ||
        lower.includes("does not exist in tenant")
      ) {
        setEmailSendError(
          "sana@epsfresno.com is not in the Fire Fly Azure tenant. In Azure Portal, set this app to multi-tenant (any organizational directory), restart the app, then try Send Now again and sign in with sana@epsfresno.com.",
        );
      } else {
        setEmailSendError(raw);
      }
      setEmailSendMessage(null);
    } finally {
      setSendingEmail(false);
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

                <div className="client-schedule-layout">
                  <aside
                    className="lookup-fields client-schedule-sidebar"
                    aria-label="Client schedule"
                  >
                    {CLIENT_SCHEDULE_ITEMS.map((item) => {
                      const businessDate = scheduleDateById.get(item.id);
                      return (
                        <button
                          key={item.id}
                          type="button"
                          className={
                            selectedSchedule.id === item.id
                              ? "lookup-field-link active"
                              : "lookup-field-link"
                          }
                          onClick={() => setSelectedScheduleId(item.id)}
                        >
                          <span className="client-schedule-link-text">
                            <span className="client-schedule-link-label">
                              {item.label}
                            </span>
                            {businessDate ? (
                              <span className="client-schedule-link-date">
                                {businessDate}
                              </span>
                            ) : null}
                          </span>
                        </button>
                      );
                    })}
                  </aside>

                  <div className="client-schedule-panel">
                    <div className="client-schedule-panel-header">
                      <h3 className="section-title">
                        {selectedSchedule.label}
                        {selectedScheduleDate
                          ? ` — ${selectedScheduleDate}`
                          : ""}
                      </h3>
                    </div>

                    {(() => {
                      const workflowTabsClassName = showRequestTypeChooser
                        ? "lookup-section-tabs bd1-workflow-tabs request-workflow-tabs"
                        : "lookup-section-tabs request-workflow-tabs";

                      const requestWorkflow = (
                        <>
                          {showBd1LikeWorkflow ? (
                            <p className="bd1-workflow-kicker">
                              {showBd7Content
                                ? "Late Action Items shows only items marked Late on BD5 for the selected request type."
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
                              BD1 action items for this request type. Mark each
                              row Completed or Late.
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
                              Items marked Late on BD5 for this category
                              (Bank and Credit Card, or Other Statements).
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
                          Subject and body are filled from Email Template{" "}
                          {requestTypeId === "other-statements" ? "2" : "1"}.
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
                              className="field-input"
                              type="text"
                              inputMode="email"
                              value={emailFrom}
                              onChange={(event) =>
                                setEmailFrom(event.target.value)
                              }
                              placeholder="email@example.com"
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
                          <div className="field field--wide">
                            <span className="field-label">Subject</span>
                            <div className="field-input field-input--readonly client-schedule-email-rich">
                              {renderAutofilledText(
                                emailSubject,
                                emailAutofillHighlights,
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
                              !emailTo.trim() ||
                              !emailSubject.trim()
                            }
                            onClick={() => {
                              // Schedule wiring comes next.
                            }}
                          >
                            Schedule
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

                      if (!showRequestTypeChooser) {
                        return requestWorkflow;
                      }

                      return (
                        <div className="bd1-request-type-block">
                          <p className="bd1-request-type-kicker">
                            Step 1 — Choose request type
                          </p>
                          <div
                            className="bd1-request-choices"
                            role="tablist"
                            aria-label={`${selectedSchedule.label} request type`}
                          >
                            {BD1_SUB_MENUS.map((item) => {
                              const isActive = requestTypeId === item.id;
                              return (
                                <button
                                  key={item.id}
                                  type="button"
                                  role="tab"
                                  aria-selected={isActive}
                                  className={
                                    isActive
                                      ? "bd1-request-choice bd1-request-choice--active"
                                      : "bd1-request-choice"
                                  }
                                  onClick={() => setRequestTypeId(item.id)}
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
                                  {activeRequestType.label}
                                </h4>
                                <p className="bd1-request-body-hint">
                                  Everything below applies to this request
                                  type only.
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
