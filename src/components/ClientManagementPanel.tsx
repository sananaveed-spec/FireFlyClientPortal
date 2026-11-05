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
import { isBdAutoSendScheduleId } from "@/lib/bd-auto-email-status";
import {
  formatBusinessDayLabel,
  getScheduleBusinessDate,
  getUpcomingScheduleBusinessDate,
} from "@/lib/business-days";
import {
  EPS_FRIDAYS_TOGGLE_ID,
  epsFridayWeekScheduleId,
  getEpsFridayWeekItemsForMonth,
  isEpsFridayWeekScheduleId,
  isEpsFridaysActive,
} from "@/lib/eps-friday-weeks";
import {
  CLIENT_MANAGEMENT_DEMO_MONTH_EVENT,
  getClientManagementNow,
  ROLL_CLIENT_MANAGEMENT_MONTH_ID,
  type RepeatedTasksStore,
} from "@/lib/repeated-tasks";
import {
  ASSIGNMENT_ROLE_LABELS,
  getClientsForManagement,
  type AssignedClient,
} from "@/lib/client-management";
import {
  formatStaffAssignmentLabel,
  resolveStaffEmail,
  userMatchKeys,
} from "@/lib/staff-assign";
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
  BD9_EXCEL_SOURCE_PREFIX,
  BD13_LATE_SOURCE_PREFIX,
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
  syncBd9ExcelRowsToActionItems,
  syncBd13LateItemsFromSources,
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
  emailTemplateLabelNumber,
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
import { normalizeEmail, type PortalUser } from "@/lib/portal-users";
import type { ResourceClientEntry } from "@/lib/resource-clients";
import {
  sendGraphMail,
  type GraphMailAttachment,
} from "@/lib/send-graph-mail";
import {
  formatBd8TotalAmount,
  getBd8AttachmentSpreadsheetStats,
  parseBd8AttachmentTable,
  sumLateActionLabelAmounts,
} from "@/lib/spreadsheet-row-count";
import {
  CLIENT_TIME_ZONE_REQUIRED_MESSAGE,
  getClientEpsFridaySendSchedule,
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

/** Fallback if no User Management user has Position = Owner. */
const SCHEDULE_EMAIL_OWNER_CC_FALLBACK = "ayesha@fireflyaccountancy.com";

/** Extra CC only for Electrical Power Systems (EPS). */
const EPS_SCHEDULE_EMAIL_CC = "thomas@epsfresno.com";

function resolveOwnerEmailFromPortalUsers(
  users: ReadonlyArray<{ email?: string; position?: string }>,
): string {
  const owner = users.find(
    (user) => user.position?.trim().toLowerCase() === "owner",
  );
  const email = owner?.email?.trim() || "";
  return email || SCHEDULE_EMAIL_OWNER_CC_FALLBACK;
}

function isEpsClientProfile(profile: {
  legalName?: string;
  operatingName?: string;
  portalLink?: string;
} | null | undefined): boolean {
  if (!profile) {
    return false;
  }
  const legal = profile.legalName?.trim().toLowerCase() || "";
  const operating = profile.operatingName?.trim().toLowerCase() || "";
  // Match "Electrical Power System" / "Systems" / short "EPS".
  if (
    /electrical power systems?\b/.test(legal) ||
    /electrical power systems?\b/.test(operating) ||
    legal === "eps" ||
    operating === "eps"
  ) {
    return true;
  }
  const portal = profile.portalLink?.trim().toLowerCase() || "";
  return /\/sites\/eps(?:\/|$|\?)/i.test(portal);
}

/** Sidebar order: BD1→…→BD18, then end-of-month by date (Last 3 → Last 2 → Last 1). */
type ClientScheduleItem = { id: string; label: string };
type ScheduleFamilyId = "documents" | "questions" | "epsFridays";

const CLIENT_SCHEDULE_ITEMS: readonly ClientScheduleItem[] = [
  { id: "bd1", label: "BD1" },
  { id: "bd5", label: "BD5" },
  { id: "bd7", label: "BD7" },
  { id: "bd8", label: "BD8" },
  { id: "bd9", label: "BD9" },
  { id: "bd10", label: "BD10" },
  { id: "bd13", label: "BD13" },
  { id: "bd15", label: "BD15" },
  { id: "bd18", label: "BD18" },
  { id: "last-3-bd", label: "Last 3 BD" },
  { id: "last-2-bd", label: "Last 2 BD" },
  { id: "last-1-bd", label: "Last 1 BD" },
];

/** BD1 chain = document request; BD8 chain = question list. */
const SCHEDULE_FAMILY_BY_ID: Record<
  string,
  { family: ScheduleFamilyId; role: "main" | "followup" }
> = {
  bd1: { family: "documents", role: "main" },
  bd5: { family: "documents", role: "followup" },
  bd7: { family: "documents", role: "followup" },
  bd8: { family: "questions", role: "main" },
  bd9: { family: "questions", role: "followup" },
  bd10: { family: "questions", role: "followup" },
  bd13: { family: "questions", role: "followup" },
};

const SCHEDULE_FAMILY_META: Record<
  ScheduleFamilyId,
  { label: string; hint: string }
> = {
  documents: { label: "Document request", hint: "BD1 → BD5 / BD7" },
  questions: { label: "Question list", hint: "BD8 → BD9 / BD10 / BD13" },
  epsFridays: {
    label: "EPS Fridays",
    hint: "Week-N Friday · 10:50 AM · Template 22",
  },
};

function scheduleFamilyForId(
  scheduleId: string,
): { family: ScheduleFamilyId; role: "main" | "followup" | null } | null {
  if (isEpsFridayWeekScheduleId(scheduleId)) {
    // No "Main" pill — date shows under the Week-N Friday label.
    return { family: "epsFridays", role: null };
  }
  return SCHEDULE_FAMILY_BY_ID[scheduleId] ?? null;
}

type ScheduleSidebarGroup = {
  family: ScheduleFamilyId | null;
  items: ClientScheduleItem[];
};

function groupScheduleItemsForSidebar(
  items: readonly ClientScheduleItem[],
): ScheduleSidebarGroup[] {
  const groups: ScheduleSidebarGroup[] = [];
  const familyGroupIndex = new Map<ScheduleFamilyId, number>();
  for (const item of items) {
    const family = scheduleFamilyForId(item.id)?.family ?? null;
    if (family == null) {
      groups.push({ family: null, items: [item] });
      continue;
    }
    const existingIndex = familyGroupIndex.get(family);
    if (existingIndex != null) {
      groups[existingIndex]?.items.push(item);
      continue;
    }
    familyGroupIndex.set(family, groups.length);
    groups.push({ family, items: [item] });
  }
  return groups;
}

const CLIENT_SCHEDULE_ORDER = new Map<string, number>(
  CLIENT_SCHEDULE_ITEMS.map((item, index) => [item.id, index]),
);

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

/** Includes EPS Fridays master toggle + Week 1–6 Friday slots for Active storage. */
const ALL_SCHEDULE_IDS = [
  ...CLIENT_SCHEDULE_ITEMS.map((item) => item.id),
  EPS_FRIDAYS_TOGGLE_ID,
  ...[1, 2, 3, 4, 5, 6].map((week) => epsFridayWeekScheduleId(week)),
];

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

/** BD9 Excel rows from BD8 attachment; BD10 uses Late from this list. */
const BD9_EXCEL_SCHEDULE_IDS = ["bd9"] as const;
/** BD13 local mirrors of BD10 late items (Completed/Late stay on BD13 only). */
const BD13_LATE_SCHEDULE_IDS = ["bd13"] as const;

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

/** BD5 Late items across both request types (what BD7 / BD9 BD7 Status Check list). */
function getBd5LateActionItemsForMonth(
  store: ClientRequestListsStore,
  clientId: string,
  monthKey: string,
): ActionWorkflowItem[] {
  const month = monthKey.trim();
  if (!clientId.trim() || !month) {
    return [];
  }
  const lateById = new Map<string, ActionWorkflowItem>();
  for (const subMenu of BD1_SUB_MENUS) {
    for (const item of getActionItemsForRequestType(
      store,
      clientId,
      subMenu.id,
    )) {
      if (item.monthKey !== month || item.status !== "late") {
        continue;
      }
      lateById.set(item.id, item);
    }
  }
  return Array.from(lateById.values()).sort((a, b) => {
    const aTime = Date.parse(a.completedAt) || 0;
    const bTime = Date.parse(b.completedAt) || 0;
    return bTime - aTime;
  });
}

/** Merge Late lists for BD10 (BD7 Status Check + BD8 Excel Status Check). */
function mergeLateActionItems(
  ...groups: ReadonlyArray<ReadonlyArray<ActionWorkflowItem>>
): ActionWorkflowItem[] {
  const byId = new Map<string, ActionWorkflowItem>();
  for (const group of groups) {
    for (const item of group) {
      if (item.status !== "late") {
        continue;
      }
      byId.set(item.id, item);
    }
  }
  return Array.from(byId.values()).sort((a, b) => {
    const aTime = Date.parse(a.completedAt) || 0;
    const bTime = Date.parse(b.completedAt) || 0;
    return bTime - aTime;
  });
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

/** Match by sidebar label number, e.g. 4 → "Template 4 (BD7)". */
function findEmailTemplateByLabelNumber(
  store: EmailTemplatesStore,
  labelNumber: number,
): EmailTemplate | null {
  const matches = store.templates.filter(
    (template) => emailTemplateLabelNumber(template) === labelNumber,
  );
  if (matches.length === 0) {
    return null;
  }
  // Prefer a custom Lookup row when label numbers collide with a seed id
  // (e.g. seed template-22 vs "Template 22 (Friday Email)").
  const custom = [...matches]
    .reverse()
    .find((template) => !/^template-\d+$/i.test(template.id));
  return custom ?? matches[matches.length - 1] ?? null;
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
          title={link}
        >
          Portal Link
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
  const { email, canEdit, role } = useAccess();
  const msalAccount = accounts[0] ?? null;
  const msalName = msalAccount?.name ?? "";
  const signedInEmail =
    getAccountEmail(msalAccount) || normalizeEmail(email) || "";

  const [resources, setResources] = useState<ResourceClientEntry[]>([]);
  const [profiles, setProfiles] = useState<ClientProfileEntry[]>([]);
  const [portalUsers, setPortalUsers] = useState<PortalUser[]>([]);
  const [portalName, setPortalName] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedScheduleId, setSelectedScheduleId] = useState<string>(
    CLIENT_SCHEDULE_ITEMS[0].id,
  );
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
  const [epsFridayAttachments, setEpsFridayAttachments] = useState<
    Array<{
      id: string;
      name: string;
      contentType: string;
      contentBytes: string;
      savedAt: string;
    }>
  >([]);
  const [uploadingEpsFridayFiles, setUploadingEpsFridayFiles] = useState(false);
  const epsFridayFileRef = useRef<HTMLInputElement | null>(null);
  const bd15FileRef = useRef<HTMLInputElement | null>(null);
  const [bd15EmailStatus, setBd15EmailStatus] = useState<{
    emailSent: boolean;
    emailSentAt: string;
    comments: Array<{ id: string; text: string; createdAt: string }>;
  }>({ emailSent: false, emailSentAt: "", comments: [] });
  const [bdAutoEmailStatus, setBdAutoEmailStatus] = useState<{
    emailSent: boolean;
    lastSentAt: string;
    fromEmail: string;
    lastError: string;
    attempts: Array<{
      at: string;
      outcome: "sent" | "error";
      message: string;
      fromEmail: string;
    }>;
  }>({
    emailSent: false,
    lastSentAt: "",
    fromEmail: "",
    lastError: "",
    attempts: [],
  });
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
  /** Skip BCC persist until saved value has been loaded for this client/BD. */
  const emailBccReadyRef = useRef(false);
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
  /** EPS Week-N Friday (sidebar only for EPS clients). */
  const showEpsFridayWeek = isEpsFridayWeekScheduleId(selectedScheduleId);
  /** Email BDs with server auto-send; Schedule disabled for everyone. */
  const scheduleOnlyEmailDay = isBdAutoSendScheduleId(selectedScheduleId);
  const showBd15Content = selectedScheduleId === "bd15";
  /** BD18 is the status check for BD15 books-closed email/files. */
  const showBd18Content = selectedScheduleId === "bd18";
  /** Last 3 / 2 / 1 BD: Template 18 team email only (portal users). */
  const showLastBdSchedule = isLastBdSchedule(selectedScheduleId);
  const showEndOfMonthEmailOnly = showLastBdSchedule;
  /** BD9/BD10 review BD8 questions (Completed/Late). Email is BD10 only. */
  const showBd9Content = selectedScheduleId === "bd9";
  const showBd10Content = selectedScheduleId === "bd10";
  /** BD13: Template 5 — same late Question List Status as BD10 + Email. */
  const showBd13Content = selectedScheduleId === "bd13";
  const showBd10OrBd13Content = showBd10Content || showBd13Content;
  /** BD1-style To/CC: POC + Owner (UM) + Group ID, readonly (also BD10/13/18). */
  const bd1LikeEmailRecipients =
    scheduleOnlyEmailDay ||
    showBd10Content ||
    showBd13Content ||
    showBd18Content;
  const showBd9LikeContent =
    showBd9Content || showBd10Content || showBd13Content;
  const showReviewActionTable = showBd5Content || showBd9LikeContent;
  /** BD1/BD8/BD15/BD18 share Request/Attachments + Email tabs. */
  const showBd1LikeRequestList =
    showBd1TypeContent ||
    showBd8Content ||
    showBd15Content ||
    showBd18Content;
  /** BD15/BD18 share Attachments + Email tabs; templates differ (11 vs 7). */
  const showBd15Email = showBd15Content || showBd18Content;
  const showBd1LikeWorkflow =
    showBd1TypeContent ||
    showBd7Content ||
    showBd8Content ||
    showBd10Content ||
    showBd13Content ||
    showBd15Email ||
    showEpsFridayWeek ||
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

  // Load saved BD8 Excel attachment for BD8 (email) and BD9/BD10 (review rows).
  // Prior-month files are pruned server-side once when the month changes.
  useEffect(() => {
    const needsBd8Attachment =
      selectedScheduleId === "bd8" ||
      selectedScheduleId === "bd9" ||
      selectedScheduleId === "bd10";
    if (!selectedId || !needsBd8Attachment) {
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
          `/api/bd8-email-attachments?clientId=${encodeURIComponent(selectedId)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}&scheduleId=bd8`,
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

  // Load EPS Week-N Friday spreadsheet attachments (multi .xlsx/.xls/.csv).
  useEffect(() => {
    if (!selectedId || !showEpsFridayWeek) {
      setEpsFridayAttachments([]);
      if (epsFridayFileRef.current) {
        epsFridayFileRef.current.value = "";
      }
      return;
    }

    let cancelled = false;
    void (async () => {
      try {
        const response = await portalFetch(
          `/api/eps-friday-attachments?clientId=${encodeURIComponent(selectedId)}&scheduleId=${encodeURIComponent(selectedScheduleId)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}`,
          { cache: "no-store" },
        );
        const data = (await response.json()) as {
          attachments?: Array<{
            id?: string;
            name?: string;
            contentType?: string;
            contentBytes?: string;
            savedAt?: string;
          }>;
          error?: string;
        };
        if (cancelled) {
          return;
        }
        if (!response.ok) {
          throw new Error(data.error || "Failed to load Friday attachments.");
        }
        setEpsFridayAttachments(
          Array.isArray(data.attachments)
            ? data.attachments
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
                  contentType:
                    typeof item.contentType === "string" && item.contentType
                      ? item.contentType
                      : "application/octet-stream",
                  contentBytes: item.contentBytes as string,
                  savedAt:
                    typeof item.savedAt === "string" && item.savedAt
                      ? item.savedAt
                      : "",
                }))
            : [],
        );
      } catch {
        if (!cancelled) {
          setEpsFridayAttachments([]);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [selectedId, selectedScheduleId, scheduleNow, showEpsFridayWeek]);

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

  // Load auto-send attempt history for email BDs.
  useEffect(() => {
    if (!selectedId || !isBdAutoSendScheduleId(selectedScheduleId)) {
      setBdAutoEmailStatus({
        emailSent: false,
        lastSentAt: "",
        fromEmail: "",
        lastError: "",
        attempts: [],
      });
      return;
    }

    let cancelled = false;
    void (async () => {
      try {
        const response = await portalFetch(
          `/api/bd-auto-emails/status?clientId=${encodeURIComponent(selectedId)}&scheduleId=${encodeURIComponent(selectedScheduleId)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}`,
          { cache: "no-store" },
        );
        const data = (await response.json()) as {
          status?: {
            emailSent?: boolean;
            lastSentAt?: string;
            fromEmail?: string;
            lastError?: string;
            attempts?: Array<{
              at?: string;
              outcome?: string;
              message?: string;
              fromEmail?: string;
            }>;
          };
          error?: string;
        };
        if (cancelled) {
          return;
        }
        if (!response.ok) {
          throw new Error(data.error || "Failed to load BD auto-email status.");
        }
        const status = data.status;
        setBdAutoEmailStatus({
          emailSent: Boolean(status?.emailSent),
          lastSentAt:
            typeof status?.lastSentAt === "string" ? status.lastSentAt : "",
          fromEmail:
            typeof status?.fromEmail === "string" ? status.fromEmail : "",
          lastError:
            typeof status?.lastError === "string" ? status.lastError : "",
          attempts: Array.isArray(status?.attempts)
            ? status.attempts
                .filter(
                  (item) =>
                    typeof item.at === "string" &&
                    item.at &&
                    (item.outcome === "sent" || item.outcome === "error"),
                )
                .map((item) => ({
                  at: item.at as string,
                  outcome: item.outcome as "sent" | "error",
                  message:
                    typeof item.message === "string" ? item.message : "",
                  fromEmail:
                    typeof item.fromEmail === "string" ? item.fromEmail : "",
                }))
            : [],
        });
      } catch {
        if (!cancelled) {
          setBdAutoEmailStatus({
            emailSent: false,
            lastSentAt: "",
            fromEmail: "",
            lastError: "",
            attempts: [],
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

  // BD5 / BD9 are status-only (no tabs). BD10/BD13 have Question List Status + Email.
  useEffect(() => {
    if (showBd5Content && workflowTab !== "action") {
      setWorkflowTab("action");
    }
  }, [showBd5Content, workflowTab]);

  useEffect(() => {
    if (showBd9Content && workflowTab !== "action") {
      setWorkflowTab("action");
    }
  }, [showBd9Content, workflowTab]);

  useEffect(() => {
    if (
      showBd10OrBd13Content &&
      workflowTab !== "action" &&
      workflowTab !== "email"
    ) {
      setWorkflowTab("action");
    }
  }, [showBd10OrBd13Content, workflowTab]);

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
      ] = await Promise.all([
        fetch("/api/resource-clients", { cache: "no-store" }),
        fetch("/api/client-profiles", { cache: "no-store" }),
        fetch("/api/portal-users", { cache: "no-store" }),
        fetch("/api/client-request-lists", { cache: "no-store" }),
        fetch("/api/client-schedule-settings", { cache: "no-store" }),
        fetch("/api/holidays", { cache: "no-store" }),
        fetch("/api/email-templates", { cache: "no-store" }),
      ]);
      const resourcesData = (await resourcesRes.json()) as ResourcesResponse;
      const profilesData = (await profilesRes.json()) as ProfilesResponse;
      const usersData = (await usersRes.json()) as UsersResponse;
      const requestListsData = await requestListsRes.json();
      const scheduleSettingsData = await scheduleSettingsRes.json();
      const holidaysData = await holidaysRes.json();
      const emailTemplatesData = await emailTemplatesRes.json();

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
    () => userMatchKeys(signedInEmail, portalName, msalName),
    [signedInEmail, portalName, msalName],
  );

  const assignedClients = useMemo(
    () =>
      getClientsForManagement(resources, profiles, aliases, {
        seeAll: role === "admin",
      }),
    [resources, profiles, aliases, role],
  );

  const selectedSupervisorEmail = useMemo(() => {
    const selectedResource =
      assignedClients.find((entry) => entry.resource.id === selectedId) ??
      assignedClients[0] ??
      null;
    if (!selectedResource) {
      return "";
    }
    return resolveStaffEmail(selectedResource.resource.supervisor, portalUsers);
  }, [assignedClients, selectedId, portalUsers]);

  const selectedSupervisorLabel = useMemo(() => {
    const selectedResource =
      assignedClients.find((entry) => entry.resource.id === selectedId) ??
      assignedClients[0] ??
      null;
    if (!selectedResource) {
      return "";
    }
    return (
      formatStaffAssignmentLabel(
        selectedResource.resource.supervisor,
        portalUsers,
      ) || selectedResource.resource.supervisor.trim()
    );
  }, [assignedClients, selectedId, portalUsers]);

  /** Schedule email From = supervisor mailbox; only that account may send. */
  const emailFrom = selectedSupervisorEmail;
  const canSendAsSupervisor =
    role !== "admin" &&
    Boolean(selectedSupervisorEmail) &&
    signedInEmail === selectedSupervisorEmail;
  const emailSendBlockedReason = !selectedSupervisorEmail
    ? "Assign a Supervisor (Primary) in Resource Management before sending."
    : role === "admin"
      ? `Admins can view schedules. Emails send only from the supervisor account (${selectedSupervisorLabel || selectedSupervisorEmail}).`
      : signedInEmail !== selectedSupervisorEmail
        ? `Only the supervisor can send. Sign in as ${selectedSupervisorLabel || selectedSupervisorEmail}.`
        : null;

  // Email body signature always follows the supervisor (From), not the viewer.
  useEffect(() => {
    let cancelled = false;
    async function loadSupervisorSignature() {
      if (!selectedSupervisorEmail) {
        setEmailSignature(createEmptyEmailSignature());
        return;
      }
      try {
        const response = await portalFetch(
          `/api/email-signature?email=${encodeURIComponent(selectedSupervisorEmail)}`,
          { cache: "no-store" },
        );
        const data = (await response.json()) as {
          signature?: EmailSignature;
          error?: string;
        };
        if (cancelled) {
          return;
        }
        if (!response.ok) {
          setEmailSignature(createEmptyEmailSignature());
          return;
        }
        setEmailSignature(normalizeEmailSignature(data.signature));
      } catch {
        if (!cancelled) {
          setEmailSignature(createEmptyEmailSignature());
        }
      }
    }
    void loadSupervisorSignature();
    return () => {
      cancelled = true;
    };
  }, [selectedSupervisorEmail]);

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
  const selectedPocEmail = selected?.profile?.pocEmail?.trim() || "";
  /** Owner email from User Management (Position = Owner). */
  const scheduleOwnerCcEmail = useMemo(
    () => resolveOwnerEmailFromPortalUsers(portalUsers),
    [portalUsers],
  );
  /** Client Profile Group ID used as CC (must look like an email). */
  const selectedGroupIdCc = useMemo(() => {
    const groupId = selected?.profile?.groupId?.trim() || "";
    return groupId.includes("@") ? groupId : "";
  }, [selected]);
  const selectedIsEpsClient = useMemo(
    () => isEpsClientProfile(selected?.profile),
    [selected],
  );
  const selectedPortalLink = selected?.profile?.portalLink?.trim() || "";
  /** Greeting name for emails — first word of Client Profile POC. */
  const selectedContactFirstName =
    selectedPointOfContact.split(/\s+/)[0]?.trim() || "";
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
  const epsFridaysActive = isEpsFridaysActive([...activeScheduleIdSet]);
  const epsFridayWeekItems = useMemo((): ClientScheduleItem[] => {
    if (!selectedIsEpsClient) {
      return [];
    }
    return getEpsFridayWeekItemsForMonth(scheduleNow).map((item) => ({
      id: item.id,
      label: item.label,
    }));
  }, [selectedIsEpsClient, scheduleNow]);
  const sortScheduleItemsByMonthDate = useCallback(
    <T extends { id: string }>(items: readonly T[]): T[] => {
      const country = selected?.profile?.country;
      const state = selected?.profile?.state ?? "";
      return [...items].sort((a, b) => {
        // EPS Fridays group always first in the sidebar.
        const epsA = isEpsFridayWeekScheduleId(a.id);
        const epsB = isEpsFridayWeekScheduleId(b.id);
        if (epsA !== epsB) {
          return epsA ? -1 : 1;
        }
        if (epsA && epsB) {
          const weekA = parseInt(a.id.replace(/\D/g, ""), 10) || 0;
          const weekB = parseInt(b.id.replace(/\D/g, ""), 10) || 0;
          return weekA - weekB;
        }
        const dateA = getUpcomingScheduleBusinessDate(
          a.id,
          holidays,
          country,
          state,
          scheduleNow,
        );
        const dateB = getUpcomingScheduleBusinessDate(
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

  const activeScheduleItems = useMemo(() => {
    const base = CLIENT_SCHEDULE_ITEMS.filter((item) =>
      activeScheduleIdSet.has(item.id),
    );
    const weeks = epsFridaysActive ? epsFridayWeekItems : [];
    return sortScheduleItemsByMonthDate([...base, ...weeks]);
  }, [
    activeScheduleIdSet,
    epsFridaysActive,
    epsFridayWeekItems,
    sortScheduleItemsByMonthDate,
  ]);
  const inactiveScheduleItems = useMemo(() => {
    const base = CLIENT_SCHEDULE_ITEMS.filter(
      (item) => !activeScheduleIdSet.has(item.id),
    );
    const weeks =
      selectedIsEpsClient && !epsFridaysActive ? epsFridayWeekItems : [];
    return sortScheduleItemsByMonthDate([...base, ...weeks]);
  }, [
    activeScheduleIdSet,
    epsFridaysActive,
    epsFridayWeekItems,
    selectedIsEpsClient,
    sortScheduleItemsByMonthDate,
  ]);
  const isSelectedScheduleActive = Boolean(
    selectedClientId &&
      (showEpsFridayWeek
        ? epsFridaysActive
        : isScheduleActiveForClient(
            scheduleSettings,
            selectedClientId,
            selectedScheduleId,
            ALL_SCHEDULE_IDS,
          )),
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
    const epsOn = isEpsFridaysActive(activeIds);
    setSelectedScheduleId((current) => {
      if (isEpsFridayWeekScheduleId(current) && epsOn) {
        return current;
      }
      return activeIds.includes(current)
        ? current
        : (activeIds[0] ?? CLIENT_SCHEDULE_ITEMS[0].id);
    });
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
        selectedScheduleId !== "bd10" &&
        selectedScheduleId !== "bd13")
    ) {
      return "";
    }
    const country = selected?.profile?.country;
    const state = selected?.profile?.state ?? "";
    // BD1 → 5th business day of the same scheduled month as BD1.
    // BD7/BD8/BD10/BD13 → same schedule date.
    let dueDate: Date | null = null;
    if (selectedScheduleId === "bd1") {
      const bd1Date = getUpcomingScheduleBusinessDate(
        "bd1",
        holidays,
        country,
        state,
        scheduleNow,
      );
      dueDate = bd1Date
        ? getScheduleBusinessDate("bd5", holidays, country, state, bd1Date)
        : null;
    } else {
      dueDate = getUpcomingScheduleBusinessDate(
        selectedScheduleId,
        holidays,
        country,
        state,
        scheduleNow,
      );
    }
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
    // Migrate legacy Last N BD (Team) sidebar ids into Last N BD.
    const legacyTeam = /^last-(\d+)-bd-team$/i.exec(selectedScheduleId.trim());
    if (legacyTeam) {
      setSelectedScheduleId(`last-${legacyTeam[1]}-bd`);
    }
  }, [selectedScheduleId]);

  useEffect(() => {
    if (
      !selectedClientId ||
      (selectedScheduleId !== "bd1" &&
        selectedScheduleId !== "bd7" &&
        selectedScheduleId !== "bd8" &&
        selectedScheduleId !== "bd10" &&
        selectedScheduleId !== "bd13" &&
        selectedScheduleId !== "bd15" &&
        selectedScheduleId !== "bd18" &&
        !showEpsFridayWeek &&
        !showLastBdSchedule)
    ) {
      return;
    }
    const template = showLastBdSchedule
      ? findEmailTemplateById(emailTemplates, "template-18")
      : showEpsFridayWeek
        ? findEmailTemplateByLabelNumber(emailTemplates, 22) ??
          findEmailTemplateById(emailTemplates, "template-22")
        : selectedScheduleId === "bd15"
          ? findEmailTemplateById(emailTemplates, "template-11")
          : selectedScheduleId === "bd18"
            ? // Sidebar label "Template 7".
              findEmailTemplateByLabelNumber(emailTemplates, 7)
            : selectedScheduleId === "bd13"
              ? // Sidebar label "Template 5 (BD13)" (id may not be template-5).
                findEmailTemplateByLabelNumber(emailTemplates, 5)
              : selectedScheduleId === "bd10"
                ? // Sidebar label "Template 6 (BD10)".
                  findEmailTemplateByLabelNumber(emailTemplates, 6)
                : selectedScheduleId === "bd8"
                  ? findEmailTemplateByLabelNumber(emailTemplates, 4)
                  : selectedScheduleId === "bd7"
                    ? findEmailTemplateByLabelNumber(emailTemplates, 3)
                    : findBd1EmailTemplate(emailTemplates, requestTypeId);
    const clientName =
      selectedOperatingName.replace(/[.:]+$/g, "").trim() ||
      selectedOperatingName.trim() ||
      "[Client Name]";
    const contactName =
      selectedContactFirstName ||
      selectedPointOfContact ||
      "[Contact First Name]";
    const portalLink = selectedPortalLink || "[Portal Link]";
    // [Period] = scheduled BD month − 1 (e.g. BD1 on 2 Nov → "October").
    const scheduleBdDate =
      getUpcomingScheduleBusinessDate(
        selectedScheduleId,
        holidays,
        selected?.profile?.country,
        selected?.profile?.state ?? "",
        scheduleNow,
      ) ?? scheduleNow;
    const periodDate = new Date(
      scheduleBdDate.getFullYear(),
      scheduleBdDate.getMonth() - 1,
      1,
    );
    const periodPortalHref = resolvePortalHref(portalLink, periodDate);
    const bd1ForPeriod = getScheduleBusinessDate(
      "bd1",
      holidays,
      selected?.profile?.country,
      selected?.profile?.state ?? "",
      scheduleBdDate,
    );
    const bd1DateLabel = bd1ForPeriod
      ? formatBusinessDayLabel(bd1ForPeriod)
      : "[BD1 Date]";
    const fridayDateLabel = showEpsFridayWeek
      ? scheduleBdDate.toLocaleDateString("en-US", {
          month: "short",
          day: "numeric",
          year: "numeric",
        })
      : "[date]";
    const dueDate = emailDeadlineLabel;
    const emailMonthKey = currentMonthKey(periodDate);
    const workingMonthKey = currentMonthKey(scheduleNow);
    const typeActionItemsForMonth = (monthKey: string) =>
      getActionItemsForRequestType(
        requestLists,
        selectedClientId,
        requestTypeId,
      ).filter((item) => item.monthKey === monthKey);
    // Prefer period month; fall back to calendar/demo month (where BD5 Late lives).
    const typeActionItemsPeriod = typeActionItemsForMonth(emailMonthKey);
    const typeActionItemsWorking = typeActionItemsForMonth(workingMonthKey);
    const typeActionItems =
      typeActionItemsPeriod.length > 0
        ? typeActionItemsPeriod
        : typeActionItemsWorking;
    const bd9ExcelEmailItemsRaw = getActionItemsForSchedules(
      requestLists,
      selectedClientId,
      BD9_EXCEL_SCHEDULE_IDS,
    );
    const bd9ExcelEmailItemsPeriod = bd9ExcelEmailItemsRaw.filter(
      (item) => item.monthKey === emailMonthKey,
    );
    const bd9ExcelEmailItemsWorking = bd9ExcelEmailItemsRaw.filter(
      (item) => item.monthKey === workingMonthKey,
    );
    const bd9ExcelEmailItems =
      bd9ExcelEmailItemsPeriod.length > 0
        ? bd9ExcelEmailItemsPeriod
        : bd9ExcelEmailItemsWorking;
    const bd9Bd7LateEmailPeriod = getBd5LateActionItemsForMonth(
      requestLists,
      selectedClientId,
      emailMonthKey,
    );
    const bd9Bd7LateEmailWorking = getBd5LateActionItemsForMonth(
      requestLists,
      selectedClientId,
      workingMonthKey,
    );
    const bd9Bd7LateEmailItems =
      bd9Bd7LateEmailPeriod.length > 0
        ? bd9Bd7LateEmailPeriod
        : bd9Bd7LateEmailWorking;
    const bd13LateEmailItemsRaw = getActionItemsForSchedules(
      requestLists,
      selectedClientId,
      BD13_LATE_SCHEDULE_IDS,
    ).filter(
      (item) =>
        item.status === "late" &&
        item.sourceRequestId.startsWith(BD13_LATE_SOURCE_PREFIX),
    );
    const bd13LateEmailItemsPeriod = bd13LateEmailItemsRaw.filter(
      (item) => item.monthKey === emailMonthKey,
    );
    const bd13LateEmailItemsWorking = bd13LateEmailItemsRaw.filter(
      (item) => item.monthKey === workingMonthKey,
    );
    const bd13LateEmailItems =
      bd13LateEmailItemsPeriod.length > 0
        ? bd13LateEmailItemsPeriod
        : bd13LateEmailItemsWorking;
    const requestListScheduleId =
      selectedScheduleId === "bd8"
        ? "bd8"
        : requestTypeScheduleId("bd1", requestTypeId);
    const allRequestListItems = getClientRequestList(
      requestLists,
      selectedClientId,
      requestListScheduleId,
    ).requestItems;
    // Prefer period-month bullets; fall back to working month, then any
    // bullets on this Request List Template (inline in the body — never attached).
    const periodRequestItems = requestItemsForMonth(
      allRequestListItems,
      emailMonthKey,
    );
    const workingRequestItems = requestItemsForMonth(
      allRequestListItems,
      workingMonthKey,
    );
    const requestListItems =
      periodRequestItems.length > 0
        ? periodRequestItems
        : workingRequestItems.length > 0
          ? workingRequestItems
          : allRequestListItems.filter((item) => item.label.trim());
    // BD1/BD8 email uses Request List Template bullets directly.
    // BD7 email uses Late Action Items (BD5 Late for the request type).
    // BD10 uses all Late items from BD9 (BD7 Status Check + BD8 Excel).
    // BD13 uses BD13-local late mirrors (Complete on BD13 does not clear BD10).
    const emailListItems =
      selectedScheduleId === "bd7"
        ? typeActionItems.filter((item) => item.status === "late")
        : selectedScheduleId === "bd10"
          ? mergeLateActionItems(
              bd9Bd7LateEmailItems,
              bd9ExcelEmailItems.filter((item) => item.status === "late"),
            )
          : selectedScheduleId === "bd13"
            ? bd13LateEmailItems
            : requestListItems;
    const bankListPlaceholder =
      "• [List of bank and credit card accounts, with last 4 digits]";
    const otherListPlaceholder =
      "• [List of loan, line of credit, lease, investment or other accounts]";
    const outstandingListPlaceholder = "• [List of outstanding items only]";
    const lateActionItemsPlaceholder = "• [Late Action Items]";
    const requestListItemsPlaceholder = "• [Request List Items]";
    const toBulletList = (items: { label: string }[]) =>
      items
        .map((item) => item.label.trim())
        .filter(Boolean)
        .map((label) => `• ${label}`)
        .join("\n");
    const actionListText =
      emailListItems.length > 0
        ? toBulletList(emailListItems)
        : selectedScheduleId === "bd7"
          ? lateActionItemsPlaceholder
          : selectedScheduleId === "bd9" || selectedScheduleId === "bd10"
            ? outstandingListPlaceholder
            : selectedScheduleId === "bd1" || selectedScheduleId === "bd8"
              ? requestListItemsPlaceholder
              : requestTypeId === "other-statements"
                ? otherListPlaceholder
                : bankListPlaceholder;
    // BD8: [Number] = filled Excel/CSV rows − 1; [Total Amount] = sum(Amount).
    // BD13: [Total Amount] = sum of Amount values on Late items.
    const bd8SheetStats =
      selectedScheduleId === "bd8" && questionListAttachment
        ? getBd8AttachmentSpreadsheetStats(
            questionListAttachment.name,
            questionListAttachment.contentType,
            questionListAttachment.contentBytes,
          )
        : null;
    const bd13LateTotalAmount =
      selectedScheduleId === "bd13"
        ? sumLateActionLabelAmounts(emailListItems)
        : null;
    const openQuestionCount =
      selectedScheduleId === "bd8"
        ? (bd8SheetStats?.dataRowCount ?? 0)
        : emailListItems.length;
    const questionCount =
      openQuestionCount > 0 ? String(openQuestionCount) : "[Number]";
    const totalAmountLabel =
      selectedScheduleId === "bd8" &&
      bd8SheetStats &&
      bd8SheetStats.totalAmount != null
        ? formatBd8TotalAmount(bd8SheetStats.totalAmount)
        : selectedScheduleId === "bd13" && bd13LateTotalAmount != null
          ? formatBd8TotalAmount(bd13LateTotalAmount)
          : "[Total Amount]";
    const questionListText =
      emailListItems.length > 0
        ? toBulletList(emailListItems)
        : "• [Question List]";

    function fillPlaceholders(text: string) {
      return applyPeriodPlaceholder(
        applyPortalLinkPlaceholders(
          text
            .replace(/\[Client Name\]/gi, clientName)
            .replace(/\[Client\]/gi, clientName)
            .replace(/\[Contact First Name\]/gi, contactName)
            .replace(/\[Due Date\]/gi, dueDate)
            .replace(/\[Deadline\]/gi, dueDate)
            .replace(/\[BD1 Date\]/gi, bd1DateLabel)
            .replace(/\[date\]/gi, fridayDateLabel)
            .replace(
              /\[Last BD Label\]/gi,
              lastBdLabel(selectedScheduleId),
            )
            .replace(/\[Number\]/gi, questionCount)
            .replace(/\[Total Amount\]/gi, totalAmountLabel)
            .replace(
              /(?:•\t*\s*)?\[Question List Items\]/gi,
              questionListText,
            )
            .replace(/(?:•\s*)?\[Question List\]/gi, questionListText)
            .replace(
              /(?:•\t*\s*)?\[Request List Items\]/gi,
              actionListText,
            )
            .replace(
              /(?:•\t*\s*)?\[Late Action Items\]/gi,
              actionListText,
            )
            .replace(
              /(?:•\t*\s*)?\[Outstanding items\]/gi,
              actionListText,
            )
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
    // BD1: Request List Template items must appear as body bullets (never as
    // an attachment / leftover [Request List Items] chip).
    if (
      selectedScheduleId === "bd1" &&
      emailListItems.length > 0 &&
      (!emailListItems.every((item) => body.includes(item.label)) ||
        /\[Request List Items\]/i.test(body) ||
        /\[List of bank and credit card accounts/i.test(body) ||
        /\[List of loan, line of credit/i.test(body))
    ) {
      const withList = body
        .replace(/(?:•\t*\s*)?\[Request List Items\]/gi, actionListText)
        .replace(
          /(?:•\t*\s*)?\[List of bank and credit card accounts, with last 4 digits\]/gi,
          actionListText,
        )
        .replace(
          /(?:•\t*\s*)?\[List of loan, line of credit, lease, investment or other accounts\]/gi,
          actionListText,
        );
      if (
        withList !== body &&
        emailListItems.every((item) => withList.includes(item.label))
      ) {
        body = withList;
      } else if (
        !emailListItems.every((item) => body.includes(item.label))
      ) {
        body = body.replace(
          /(need the statements(?: for the accounts)? below:|statements below for \[Period\] so we can reconcile[^\n]*:)\s*/i,
          `$1\n${actionListText}\n\n`,
        );
        if (!emailListItems.every((item) => body.includes(item.label))) {
          body = `${body.trimEnd()}\n\n${actionListText}`;
        }
      } else {
        body = withList;
      }
    }
    // BD8/BD10/BD13: ensure the question list bullets appear in the email body.
    const isQuestionEmailDay =
      selectedScheduleId === "bd8" ||
      selectedScheduleId === "bd10" ||
      selectedScheduleId === "bd13";
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
        .replace(
          /(?:•\t*\s*)?\[Question List Items\]/i,
          questionListText,
        )
        .replace(/(?:•\s*)?\[Question List\]/i, questionListText);
      if (withList !== body) {
        body = withList;
      } else {
        body = `${body.trimEnd()}\n\n${questionListText}`;
      }
    }
    setEmailSubject(subject);
    setEmailBody(body);
    setEmailPortalHref(periodPortalHref);
    // Do not bold auto-filled placeholders in preview or outbound mail.
    setEmailAutofillHighlights([]);
  }, [
    selectedClientId,
    selectedScheduleId,
    showLastBdSchedule,
    showEpsFridayWeek,
    requestTypeId,
    selectedOperatingName,
    selectedPointOfContact,
    selectedContactFirstName,
    selectedPortalLink,
    emailDeadlineLabel,
    requestLists,
    emailTemplates,
    scheduleNow,
    holidays,
    selected?.profile?.country,
    selected?.profile?.state,
    questionListAttachment,
  ]);

  useEffect(() => {
    if (
      !selectedClientId ||
      (selectedScheduleId !== "bd1" &&
        selectedScheduleId !== "bd7" &&
        selectedScheduleId !== "bd8" &&
        selectedScheduleId !== "bd10" &&
        selectedScheduleId !== "bd13" &&
        selectedScheduleId !== "bd15" &&
        selectedScheduleId !== "bd18" &&
        !showEpsFridayWeek &&
        !showLastBdSchedule)
    ) {
      return;
    }
    // From is the supervisor mailbox. BD1 To = client POC email (schedule only).
    // Last N BD team reminder: To = supervisor, BCC = all portal users.
    if (showLastBdSchedule) {
      const teamEmails = portalUsers
        .map((user) => user.email.trim().toLowerCase())
        .filter(Boolean);
      const sender = emailFrom.trim().toLowerCase();
      const bccList = teamEmails.filter((item) => item !== sender);
      setEmailTo(sender || teamEmails[0] || "");
      setEmailCc("");
      setEmailBcc(bccList.join(", "));
    } else if (
      showBd1TypeContent ||
      showBd7Content ||
      showBd8Content ||
      showBd10Content ||
      showBd13Content ||
      showBd15Content ||
      showBd18Content ||
      showEpsFridayWeek
    ) {
      setEmailTo(selectedPocEmail);
      const toEmail = selectedPocEmail.trim().toLowerCase();
      const fromEmail = emailFrom.trim().toLowerCase();
      const ccEmails = [
        scheduleOwnerCcEmail,
        selectedGroupIdCc,
        selectedIsEpsClient || showEpsFridayWeek ? EPS_SCHEDULE_EMAIL_CC : "",
      ]
        .map((email) => email.trim().toLowerCase())
        .filter(
          (email) =>
            Boolean(email) && email !== toEmail && email !== fromEmail,
        );
      setEmailCc([...new Set(ccEmails)].join(", "));
      // BCC is user-editable and persisted — do not clear here.
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
    showLastBdSchedule,
    showBd1TypeContent,
    showBd7Content,
    showBd8Content,
    showBd10Content,
    showBd13Content,
    showBd15Content,
    showBd18Content,
    showEpsFridayWeek,
    requestTypeId,
    portalUsers,
    emailFrom,
    selectedPocEmail,
    scheduleOwnerCcEmail,
    selectedGroupIdCc,
    selectedIsEpsClient,
  ]);

  // Load persisted BCC for BD email tabs (survives refresh / re-open).
  useEffect(() => {
    if (
      !selectedId ||
      showLastBdSchedule ||
      !(
        showBd1TypeContent ||
        showBd7Content ||
        showBd8Content ||
        showBd10Content ||
        showBd13Content ||
        showBd15Content ||
        showBd18Content ||
        showEpsFridayWeek
      )
    ) {
      emailBccReadyRef.current = false;
      return;
    }

    let cancelled = false;
    emailBccReadyRef.current = false;
    setEmailBcc("");
    void (async () => {
      try {
        const response = await portalFetch(
          `/api/bd-email-bcc?clientId=${encodeURIComponent(selectedId)}&scheduleId=${encodeURIComponent(selectedScheduleId)}`,
          { cache: "no-store" },
        );
        const data = (await response.json()) as {
          bcc?: string;
          error?: string;
        };
        if (cancelled) {
          return;
        }
        setEmailBcc(
          response.ok && typeof data.bcc === "string" ? data.bcc : "",
        );
      } catch {
        if (!cancelled) {
          setEmailBcc("");
        }
      } finally {
        if (!cancelled) {
          emailBccReadyRef.current = true;
        }
      }
    })();

    return () => {
      cancelled = true;
      emailBccReadyRef.current = false;
    };
  }, [
    selectedId,
    selectedScheduleId,
    showLastBdSchedule,
    showBd1TypeContent,
    showBd7Content,
    showBd8Content,
    showBd10Content,
    showBd13Content,
    showBd15Content,
    showBd18Content,
    showEpsFridayWeek,
  ]);

  // Persist BCC after the user edits it (debounced).
  useEffect(() => {
    if (
      !selectedId ||
      !emailBccReadyRef.current ||
      showLastBdSchedule ||
      !(
        showBd1TypeContent ||
        showBd7Content ||
        showBd8Content ||
        showBd10Content ||
        showBd13Content ||
        showBd15Content ||
        showBd18Content ||
        showEpsFridayWeek
      )
    ) {
      return;
    }

    const timer = window.setTimeout(() => {
      if (!emailBccReadyRef.current) {
        return;
      }
      void (async () => {
        try {
          await portalFetch("/api/bd-email-bcc", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              clientId: selectedId,
              scheduleId: selectedScheduleId,
              bcc: emailBcc,
            }),
          });
        } catch {
          // Best-effort save; field still keeps local value.
        }
      })();
    }, 500);

    return () => {
      window.clearTimeout(timer);
    };
  }, [
    emailBcc,
    selectedId,
    selectedScheduleId,
    showLastBdSchedule,
    showBd1TypeContent,
    showBd7Content,
    showBd8Content,
    showBd10Content,
    showBd13Content,
    showBd15Content,
    showBd18Content,
    showEpsFridayWeek,
  ]);

  const selectedSchedule =
    CLIENT_SCHEDULE_ITEMS.find((item) => item.id === selectedScheduleId) ??
    epsFridayWeekItems.find((item) => item.id === selectedScheduleId) ??
    CLIENT_SCHEDULE_ITEMS[0];

  const selectedScheduleBusinessDate = useMemo(() => {
    if (!selected) {
      return null;
    }
    return getUpcomingScheduleBusinessDate(
      selectedScheduleId,
      holidays,
      selected.profile?.country,
      selected.profile?.state ?? "",
      scheduleNow,
    );
  }, [selected, selectedScheduleId, holidays, scheduleNow]);

  // BD date at 9:00 / EPS Friday at 10:50 client TZ → Pakistan (read-only).
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
      showEpsFridayWeek
        ? getClientEpsFridaySendSchedule(
            selectedScheduleBusinessDate,
            selected?.profile?.timeZone,
          )
        : getClientNineAmSchedule(
            selectedScheduleBusinessDate,
            selected?.profile?.timeZone,
          ),
    [
      showEpsFridayWeek,
      selectedScheduleBusinessDate,
      selected?.profile?.timeZone,
    ],
  );
  const emailSendAtClientLabel = defaultEmailSendSchedule?.clientLabel ?? "";
  const emailSendAtPakistanLabel = defaultEmailSendSchedule?.userLabel ?? "";
  const effectiveEmailSendAt =
    emailSendAt || defaultEmailSendSchedule?.userDateTimeLocal || "";

  const scheduleDateById = useMemo(() => {
    const map = new Map<string, string>();
    if (!selected) {
      return map;
    }
    const country = selected.profile?.country;
    const state = selected.profile?.state ?? "";
    const items = [...CLIENT_SCHEDULE_ITEMS, ...epsFridayWeekItems];
    for (const item of items) {
      const date = getUpcomingScheduleBusinessDate(
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
  }, [selected, holidays, scheduleNow, epsFridayWeekItems]);

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
  const bd9ExcelActionItems =
    showBd9LikeContent && selected
      ? getActionItemsForSchedules(
          requestLists,
          selected.resource.id,
          BD9_EXCEL_SCHEDULE_IDS,
        )
      : [];
  const bd9ExcelTable = useMemo(() => {
    if (
      !showBd9LikeContent ||
      !questionListAttachment?.contentBytes ||
      !questionListAttachment.name
    ) {
      return null;
    }
    return parseBd8AttachmentTable(
      questionListAttachment.name,
      questionListAttachment.contentType,
      questionListAttachment.contentBytes,
    );
  }, [showBd9LikeContent, questionListAttachment]);
  const monthScopedActionItems = (
    items: ActionWorkflowItem[],
  ): ActionWorkflowItem[] =>
    items.filter((item) => item.monthKey === scheduleMonthKey);
  /** BD9 BD7 Status Check (+ BD10/BD13 upstream): BD5 Late items for the month. */
  const bd9Bd7StatusLateItems = useMemo(() => {
    if ((!showBd9Content && !showBd10OrBd13Content) || !selected) {
      return [];
    }
    return getBd5LateActionItemsForMonth(
      requestLists,
      selected.resource.id,
      scheduleMonthKey,
    );
  }, [
    showBd9Content,
    showBd10OrBd13Content,
    selected,
    requestLists,
    scheduleMonthKey,
  ]);
  /** Upstream late rows that BD10 shows (and BD13 mirrors locally). */
  const bd10UpstreamLateItems = useMemo(
    () =>
      mergeLateActionItems(
        bd9Bd7StatusLateItems,
        monthScopedActionItems(
          bd9ExcelActionItems.filter((item) => item.status === "late"),
        ),
      ),
    [
      bd9Bd7StatusLateItems,
      bd9ExcelActionItems,
      scheduleMonthKey,
    ],
  );
  const bd13LocalActionItems =
    showBd13Content && selected
      ? getActionItemsForSchedules(
          requestLists,
          selected.resource.id,
          BD13_LATE_SCHEDULE_IDS,
        )
      : [];
  const scopedActionItems = monthScopedActionItems(
    showBd13Content
      ? // BD13 uses local mirrors only — Complete/Late do not change BD9/BD10.
        bd13LocalActionItems.filter((item) =>
          item.sourceRequestId.startsWith(BD13_LATE_SOURCE_PREFIX),
        )
      : showBd10Content
        ? bd10UpstreamLateItems
        : showBd9Content
          ? bd9ExcelActionItems.filter((item) =>
              item.sourceRequestId.startsWith(BD9_EXCEL_SOURCE_PREFIX),
            )
          : showRequestTypeChooser && selected
            ? getActionItemsForRequestType(
                requestLists,
                selected.resource.id,
                requestTypeId,
              )
            : sharedActionItems,
  );

  // Sync BD8 Excel rows onto the BD9 action list for Completed/Late review.
  useEffect(() => {
    if (
      !showBd9Content ||
      !selected ||
      !questionListAttachment ||
      !bd9ExcelTable
    ) {
      return;
    }
    const list = getClientRequestList(
      requestLists,
      selected.resource.id,
      "bd9",
    );
    const synced = syncBd9ExcelRowsToActionItems(
      list,
      bd9ExcelTable.rows,
      scheduleMonthKey,
    );
    if (synced === list) {
      return;
    }
    const before = JSON.stringify(list.actionItems);
    const after = JSON.stringify(synced.actionItems);
    if (before === after) {
      return;
    }
    void persistList(synced);
    // persistList is stable enough for this sync; avoid re-running on identity churn.
    // eslint-disable-next-line react-hooks/exhaustive-deps -- intentional
  }, [
    showBd9Content,
    selected?.resource.id,
    questionListAttachment,
    bd9ExcelTable,
    scheduleMonthKey,
    requestLists,
  ]);

  // Mirror BD10 late items onto BD13 so Complete/Late stay local to BD13.
  useEffect(() => {
    if (!showBd13Content || !selected) {
      return;
    }
    const list = getClientRequestList(
      requestLists,
      selected.resource.id,
      "bd13",
    );
    const synced = syncBd13LateItemsFromSources(
      list,
      bd10UpstreamLateItems.map((item) => ({
        id: item.id,
        label: item.label,
        completedAt: item.completedAt,
      })),
      scheduleMonthKey,
    );
    if (synced === list) {
      return;
    }
    const before = JSON.stringify(list.actionItems);
    const after = JSON.stringify(synced.actionItems);
    if (before === after) {
      return;
    }
    void persistList(synced);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- intentional
  }, [
    showBd13Content,
    selected?.resource.id,
    bd10UpstreamLateItems,
    scheduleMonthKey,
    requestLists,
  ]);
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
    // BD13: keep Late and Completed visible after marking Complete.
    if (showBd13Content) {
      const late = scopedActionItems.filter((item) => item.status === "late");
      const completed = scopedActionItems
        .filter((item) => item.status === "completed")
        .sort((a, b) => {
          const aTime = Date.parse(a.completedAt) || 0;
          const bTime = Date.parse(b.completedAt) || 0;
          return bTime - aTime;
        });
      return [...late, ...completed];
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
  }, [scopedActionItems, showBd10Content, showBd13Content]);
  const bd7EmailDisabled =
    showBd7Content && bd7LateItems.length === 0;
  const bd8EmailDisabled =
    showBd8Content && !questionListAttachment;
  const bd10EmailDisabled =
    showBd10Content && bd5ActionTableItems.length === 0;
  const bd13EmailDisabled =
    showBd13Content &&
    !scopedActionItems.some((item) => item.status === "late");
  const bd10LateStatusHint = useMemo(() => {
    if (showBd13Content) {
      return "BD10 Late appear here. Completed on BD13 stays on BD13 only.";
    }
    if (!showBd10Content) {
      return "";
    }
    const hasBd7Late = bd9Bd7StatusLateItems.length > 0;
    const hasBd8Late = monthScopedActionItems(bd9ExcelActionItems).some(
      (item) =>
        item.status === "late" &&
        item.sourceRequestId.startsWith(BD9_EXCEL_SOURCE_PREFIX),
    );
    if (!hasBd7Late && !hasBd8Late) {
      return "No late status found for BD7 and BD8.";
    }
    if (hasBd7Late && hasBd8Late) {
      return "BD7 and BD8 late status automatically appear here.";
    }
    if (hasBd7Late) {
      return "BD7 late status automatically appears here. No late status found for BD8.";
    }
    return "BD8 late status automatically appears here. No late status found for BD7.";
  }, [
    showBd13Content,
    showBd10Content,
    bd9Bd7StatusLateItems,
    bd9ExcelActionItems,
    scheduleMonthKey,
  ]);
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
    bd8EmailDisabled ||
    bd10EmailDisabled ||
    bd13EmailDisabled ||
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
      : bd8EmailDisabled
        ? "Add one Excel/CSV attachment on Question List Template before scheduling."
        : bd10EmailDisabled
          ? "No late BD9 items found."
          : bd13EmailDisabled
            ? "No late BD10 question list items found."
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

    // One EPS Fridays star toggles all Week-N Friday rows.
    const toggleId = isEpsFridayWeekScheduleId(scheduleId)
      ? EPS_FRIDAYS_TOGGLE_ID
      : scheduleId;
    const nextStore = toggleClientScheduleActive(
      scheduleSettings,
      selected.resource.id,
      toggleId,
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
          scheduleId: "bd8",
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
      event.target.value = "";
    }
  }

  async function handleClearQuestionListAttachment() {
    if (!selected) {
      return;
    }
    if (!requireActiveSchedule("clear the Excel attachment")) {
      return;
    }
    setUploadingQuestionList(true);
    setError(null);
    try {
      const response = await portalFetch(
        `/api/bd8-email-attachments?clientId=${encodeURIComponent(selected.resource.id)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}&scheduleId=bd8`,
        { method: "DELETE" },
      );
      const data = (await response.json()) as { error?: string };
      if (!response.ok) {
        throw new Error(data.error || "Failed to clear Excel attachment.");
      }
      setQuestionListAttachment(null);
      if (questionListFileRef.current) {
        questionListFileRef.current.value = "";
      }
    } catch (clearError) {
      setError(
        clearError instanceof Error
          ? clearError.message
          : "Failed to clear Excel attachment.",
      );
    } finally {
      setUploadingQuestionList(false);
    }
  }

  async function handleEpsFridayFileUpload(
    event: ChangeEvent<HTMLInputElement>,
  ) {
    const files = Array.from(event.target.files ?? []);
    if (files.length === 0 || !selected || !showEpsFridayWeek) {
      return;
    }
    if (!requireActiveSchedule("attach Friday files")) {
      event.target.value = "";
      return;
    }

    const allowed = files.filter((file) => {
      const lower = file.name.toLowerCase();
      return (
        lower.endsWith(".xlsx") ||
        lower.endsWith(".xls") ||
        lower.endsWith(".csv")
      );
    });
    if (allowed.length === 0) {
      setError("Choose .xlsx, .xls, or .csv files only.");
      event.target.value = "";
      return;
    }

    setUploadingEpsFridayFiles(true);
    setError(null);
    try {
      const monthKey = currentMonthKey(scheduleNow);
      const payload = [];
      for (const file of allowed) {
        const buffer = await file.arrayBuffer();
        payload.push({
          clientId: selected.resource.id,
          scheduleId: selectedScheduleId,
          monthKey,
          name: file.name,
          contentType: mimeTypeForExcelFile(file),
          contentBytes: arrayBufferToBase64(buffer),
        });
      }
      const response = await portalFetch("/api/eps-friday-attachments", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ attachments: payload }),
      });
      const data = (await response.json()) as {
        attachments?: Array<{
          id?: string;
          name?: string;
          contentType?: string;
          contentBytes?: string;
          savedAt?: string;
        }>;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to save Friday attachments.");
      }
      setEpsFridayAttachments(
        Array.isArray(data.attachments)
          ? data.attachments
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
                contentType:
                  typeof item.contentType === "string" && item.contentType
                    ? item.contentType
                    : "application/octet-stream",
                contentBytes: item.contentBytes as string,
                savedAt:
                  typeof item.savedAt === "string" && item.savedAt
                    ? item.savedAt
                    : "",
              }))
          : [],
      );
    } catch (uploadError) {
      setError(
        uploadError instanceof Error
          ? uploadError.message
          : "Failed to save Friday attachments.",
      );
    } finally {
      setUploadingEpsFridayFiles(false);
      event.target.value = "";
    }
  }

  async function handleRemoveEpsFridayAttachment(attachmentId: string) {
    if (!selected || !showEpsFridayWeek) {
      return;
    }
    if (!requireActiveSchedule("remove Friday files")) {
      return;
    }
    setUploadingEpsFridayFiles(true);
    setError(null);
    try {
      const response = await portalFetch(
        `/api/eps-friday-attachments?clientId=${encodeURIComponent(selected.resource.id)}&scheduleId=${encodeURIComponent(selectedScheduleId)}&monthKey=${encodeURIComponent(currentMonthKey(scheduleNow))}&attachmentId=${encodeURIComponent(attachmentId)}`,
        { method: "DELETE" },
      );
      const data = (await response.json()) as {
        attachments?: Array<{
          id?: string;
          name?: string;
          contentType?: string;
          contentBytes?: string;
          savedAt?: string;
        }>;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to remove attachment.");
      }
      setEpsFridayAttachments(
        Array.isArray(data.attachments)
          ? data.attachments
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
                contentType:
                  typeof item.contentType === "string" && item.contentType
                    ? item.contentType
                    : "application/octet-stream",
                contentBytes: item.contentBytes as string,
                savedAt:
                  typeof item.savedAt === "string" && item.savedAt
                    ? item.savedAt
                    : "",
              }))
          : [],
      );
    } catch (removeError) {
      setError(
        removeError instanceof Error
          ? removeError.message
          : "Failed to remove attachment.",
      );
    } finally {
      setUploadingEpsFridayFiles(false);
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
    // Allow Completed/Late on already-resolved rows (e.g. BD13 late review).
    await persistList(setActionItemStatus(owner, itemId, status));
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
    if (scheduleOnlyEmailDay) {
      const scheduleOnlyLabel = showBd8Content
        ? "BD8"
        : showBd7Content
          ? "BD7"
          : "BD1";
      setEmailSendError(
        `${scheduleOnlyLabel} is schedule-only. Use Schedule for the upcoming ${scheduleOnlyLabel} send time.`,
      );
      setEmailSendMessage(null);
      return;
    }
    if (!canSendAsSupervisor) {
      setEmailSendError(
        emailSendBlockedReason ||
          "Only the assigned supervisor can send this email.",
      );
      setEmailSendMessage(null);
      return;
    }
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
    if (scheduleOnlyEmailDay) {
      const scheduleOnlyLabel = showEpsFridayWeek
        ? selectedSchedule.label
        : showBd8Content
          ? "BD8"
          : showBd7Content
            ? "BD7"
            : "BD1";
      const sendTimeLabel = showEpsFridayWeek
        ? "10:50 am client time"
        : "9:00 am client time";
      setEmailSendError(
        `${scheduleOnlyLabel} is sent automatically by the server at ${sendTimeLabel}. Schedule is disabled.`,
      );
      setEmailSendMessage(null);
      return;
    }
    if (!canSendAsSupervisor) {
      setEmailSendError(
        emailSendBlockedReason ||
          "Only the assigned supervisor can schedule this email.",
      );
      setEmailSendMessage(null);
      return;
    }
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
      if (
        selectedScheduleId === "bd1" ||
        selectedScheduleId === "bd7" ||
        selectedScheduleId === "bd8"
      ) {
        try {
          const markResponse = await portalFetch("/api/bd-auto-emails/status", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              action: "mark-sent",
              clientId: selected?.resource.id,
              scheduleId: selectedScheduleId,
              monthKey: currentMonthKey(scheduleNow),
              lastSentAt: sendAt.toISOString(),
              fromEmail: emailFrom,
              source: "manual-schedule",
            }),
          });
          const markData = (await markResponse.json()) as {
            status?: {
              emailSent?: boolean;
              lastSentAt?: string;
              fromEmail?: string;
              lastError?: string;
            };
          };
          if (markResponse.ok && markData.status) {
            setBdAutoEmailStatus({
              emailSent: Boolean(markData.status.emailSent),
              lastSentAt: markData.status.lastSentAt ?? sendAt.toISOString(),
              fromEmail: markData.status.fromEmail ?? emailFrom,
              lastError: markData.status.lastError ?? "",
              attempts: [],
            });
          }
        } catch {
          // Schedule already queued in Outlook; status write is best-effort.
        }
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
          {role === "admin"
            ? "No clients in Resource Allocation yet."
            : "No clients are assigned to you in Resource Allocation yet."}
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
                        groupScheduleItemsForSidebar(activeScheduleItems).map(
                          (group, groupIndex) => {
                            const familyMeta = group.family
                              ? SCHEDULE_FAMILY_META[group.family]
                              : null;
                            const rows = group.items.map((item) => {
                              const businessDate = scheduleDateById.get(
                                item.id,
                              );
                              const isSelected =
                                selectedSchedule.id === item.id;
                              const role =
                                scheduleFamilyForId(item.id)?.role ?? null;
                              const rowClass = [
                                "schedule-row",
                                "schedule-row--on",
                                isSelected ? "schedule-row--selected" : "",
                                role === "main" ? "schedule-row--main" : "",
                                role === "followup"
                                  ? "schedule-row--followup"
                                  : "",
                              ]
                                .filter(Boolean)
                                .join(" ");
                              return (
                                <div key={item.id} className={rowClass}>
                                  <button
                                    type="button"
                                    className="schedule-star schedule-star--filled"
                                    aria-label={`Deactivate ${item.label}`}
                                    title="Move to Inactive"
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
                                    ★
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
                                      {role === "main" ? (
                                        <span className="schedule-role-pill">
                                          Main
                                        </span>
                                      ) : null}
                                    </span>
                                    {businessDate ? (
                                      <span className="client-schedule-link-date">
                                        {businessDate}
                                      </span>
                                    ) : null}
                                  </button>
                                </div>
                              );
                            });
                            if (!group.family || !familyMeta) {
                              return (
                                <div
                                  key={`solo-${group.items[0]?.id ?? groupIndex}`}
                                  className="schedule-family schedule-family--solo"
                                >
                                  {rows}
                                </div>
                              );
                            }
                            return (
                              <div
                                key={group.family}
                                className={`schedule-family schedule-family--${group.family}`}
                              >
                                <div className="schedule-family-heading">
                                  <p className="schedule-family-label">
                                    {familyMeta.label}
                                  </p>
                                  <p className="schedule-family-hint">
                                    {familyMeta.hint}
                                  </p>
                                </div>
                                {rows}
                              </div>
                            );
                          },
                        )
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
                          {groupScheduleItemsForSidebar(
                            inactiveScheduleItems,
                          ).map((group, groupIndex) => {
                            const familyMeta = group.family
                              ? SCHEDULE_FAMILY_META[group.family]
                              : null;
                            const rows = group.items.map((item) => {
                              const businessDate = scheduleDateById.get(
                                item.id,
                              );
                              const isSelected =
                                selectedSchedule.id === item.id;
                              const role =
                                scheduleFamilyForId(item.id)?.role ?? null;
                              const rowClass = [
                                "schedule-row",
                                "schedule-row--off",
                                isSelected ? "schedule-row--selected" : "",
                                role === "main" ? "schedule-row--main" : "",
                                role === "followup"
                                  ? "schedule-row--followup"
                                  : "",
                              ]
                                .filter(Boolean)
                                .join(" ");
                              return (
                                <div key={item.id} className={rowClass}>
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
                                      {role === "main" ? (
                                        <span className="schedule-role-pill">
                                          Main
                                        </span>
                                      ) : null}
                                    </span>
                                    {businessDate ? (
                                      <span className="client-schedule-link-date">
                                        {businessDate}
                                      </span>
                                    ) : null}
                                  </button>
                                </div>
                              );
                            });
                            if (!group.family || !familyMeta) {
                              return (
                                <div
                                  key={`solo-off-${group.items[0]?.id ?? groupIndex}`}
                                  className="schedule-family schedule-family--solo"
                                >
                                  {rows}
                                </div>
                              );
                            }
                            return (
                              <div
                                key={`${group.family}-off`}
                                className={`schedule-family schedule-family--${group.family}`}
                              >
                                <div className="schedule-family-heading">
                                  <p className="schedule-family-label">
                                    {familyMeta.label}
                                  </p>
                                  <p className="schedule-family-hint">
                                    {familyMeta.hint}
                                  </p>
                                </div>
                                {rows}
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
                          {!showBd5Content &&
                          !showBd9Content &&
                          !showEndOfMonthEmailOnly ? (
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
                      {showBd10OrBd13Content ||
                      showBd7Content ||
                      (!showBd1LikeRequestList &&
                        !showBd5Content &&
                        !showBd9Content) ? (
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
                        {showBd10OrBd13Content
                          ? "Question List Status"
                          : showBd7Content
                            ? "Late Action Items"
                            : showEpsFridayWeek
                              ? "Attachments"
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
                      {!showBd1LikeWorkflow &&
                      !showBd5Content &&
                      !showBd9Content ? (
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
                          {showBd10OrBd13Content && bd10LateStatusHint ? (
                            <p className="bd1-workflow-kicker">
                              {bd10LateStatusHint}
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
                            ? "Add questions as text below. One Excel/CSV file is required as the email attachment ([Number] = filled rows − 1; [Total Amount] = sum of Amount column; Email stays disabled until attached)."
                            : "Add items here. The Email tab uses these bullets directly for this request type."}
                        </p>
                        {showBd8Content ? (
                          <div className="question-list-upload">
                            <label className="field field--wide">
                              <span className="field-label">
                                Email attachment (one file)
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
                            <div className="bd5-action-toolbar">
                              <p className="lookup-hint">
                                {uploadingQuestionList
                                  ? "Loading…"
                                  : questionListAttachment
                                    ? `Attached: ${questionListAttachment.name}`
                                    : "No attachment yet. Add .xlsx, .xls, or .csv to enable Email."}
                              </p>
                              {questionListAttachment ? (
                                <button
                                  type="button"
                                  className="button secondary"
                                  disabled={
                                    uploadingQuestionList ||
                                    savingRequestList ||
                                    !isSelectedScheduleActive
                                  }
                                  onClick={() =>
                                    void handleClearQuestionListAttachment()
                                  }
                                >
                                  Clear attachment
                                </button>
                              ) : null}
                            </div>
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
                        {showBd9Content ? (
                          <>
                            <section className="bd9-status-section">
                              <h3 className="bd9-status-section-title">
                                BD7 Status Check
                              </h3>
                              <p className="lookup-hint">
                                List items marked Late on BD5 (same items that
                                appear on BD7).
                              </p>
                              {bd9Bd7StatusLateItems.length === 0 ? (
                                <p className="lookup-hint">
                                  No late action items from BD7 found.
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
                                      {bd9Bd7StatusLateItems.map((item) => {
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
                                            <td>
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
                                            </td>
                                          </tr>
                                        );
                                      })}
                                    </tbody>
                                  </table>
                                </div>
                              )}
                            </section>

                            <section className="bd9-status-section">
                              <h3 className="bd9-status-section-title">
                                BD8 Status Check
                              </h3>
                              <p className="lookup-hint">
                                All rows from the BD8 Excel/CSV attachment,
                                whether Open, Completed, or Late. Mark Open
                                rows Completed or Late — Late rows appear on
                                BD10.
                                {!questionListAttachment
                                  ? " Attach a file on BD8 Question List Template first."
                                  : ""}
                              </p>
                              {bd5ActionTableItems.length === 0 ? (
                                <p className="lookup-hint">
                                  {questionListAttachment
                                    ? "No data rows found in the BD8 Excel/CSV attachment."
                                    : "No BD8 Excel/CSV attachment for this month yet."}
                                </p>
                              ) : (
                                <div className="table-wrap bd5-action-table-wrap">
                                  <table className="projects-table bd5-action-table">
                                    <thead>
                                      <tr>
                                        {bd9ExcelTable &&
                                        bd9ExcelTable.headers.length > 0
                                          ? bd9ExcelTable.headers.map(
                                              (header, index) => (
                                                <th
                                                  key={`bd9-h-${index}-${header}`}
                                                  scope="col"
                                                >
                                                  {header.trim() ||
                                                    `Column ${index + 1}`}
                                                </th>
                                              ),
                                            )
                                          : (
                                            <th scope="col">List</th>
                                          )}
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
                                        const isOpen =
                                          item.status === "action";
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
                                        const excelFingerprint =
                                          item.sourceRequestId.startsWith(
                                            BD9_EXCEL_SOURCE_PREFIX,
                                          )
                                            ? item.sourceRequestId.slice(
                                                BD9_EXCEL_SOURCE_PREFIX.length,
                                              )
                                            : "";
                                        const excelRow = bd9ExcelTable
                                          ? bd9ExcelTable.rows.find(
                                              (row) =>
                                                row.fingerprint ===
                                                excelFingerprint,
                                            )
                                          : undefined;
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
                                            {bd9ExcelTable &&
                                            bd9ExcelTable.headers.length >
                                              0 ? (
                                              bd9ExcelTable.headers.map(
                                                (_header, index) => (
                                                  <td
                                                    key={`${item.id}-c-${index}`}
                                                    className="bd5-action-list-cell"
                                                  >
                                                    {excelRow?.cells[index] ??
                                                      "—"}
                                                  </td>
                                                ),
                                              )
                                            ) : (
                                              <td className="bd5-action-list-cell">
                                                {item.label}
                                              </td>
                                            )}
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
                                                    disabled={
                                                      savingRequestList
                                                    }
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
                                                    disabled={
                                                      savingRequestList
                                                    }
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
                            </section>
                          </>
                        ) : showReviewActionTable ? (
                          <>
                            {showBd5Content ? (
                              <>
                                <div className="bd5-action-toolbar">
                                  <p className="lookup-hint">
                                    SharePoint auto-check is paused until search
                                    logic is defined for this client’s folder
                                    files.
                                  </p>
                                  <button
                                    type="button"
                                    className="button secondary-solid"
                                    disabled
                                    title="Search logic needed before SharePoint check can run"
                                  >
                                    Search Logic Needed
                                  </button>
                                </div>
                                {sharePointCheckMessage ? (
                                  <p className="lookup-hint bd5-sharepoint-result">
                                    {sharePointCheckMessage}
                                  </p>
                                ) : null}
                              </>
                            ) : null}
                            {bd5ActionTableItems.length === 0 ? (
                              <p className="lookup-hint">
                                {showBd10Content
                                  ? "No late BD9 items found."
                                  : showBd13Content
                                    ? "No late or completed BD13 items found."
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
                                      {showBd10Content ? null : (
                                        <th scope="col">Action</th>
                                      )}
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
                                          {showBd10Content ? null : (
                                          <td>
                                            {showBd13Content || isOpen ? (
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
                                          )}
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            )}
                          </>
                        ) : showEpsFridayWeek ? (
                          <div className="question-list-upload">
                            <p className="lookup-hint">
                              Upload one or more .xlsx, .xls, or .csv files for
                              this Week Friday email. Files are attached on
                              auto-send at 10:50 am client time.
                            </p>
                            <label className="field field--wide">
                              <span className="field-label">
                                Spreadsheet attachments
                              </span>
                              <input
                                ref={epsFridayFileRef}
                                className="field-input"
                                type="file"
                                multiple
                                accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv"
                                disabled={
                                  uploadingEpsFridayFiles ||
                                  !isSelectedScheduleActive
                                }
                                onChange={(event) =>
                                  void handleEpsFridayFileUpload(event)
                                }
                              />
                            </label>
                            <div className="bd5-action-toolbar">
                              <p className="lookup-hint">
                                {uploadingEpsFridayFiles
                                  ? "Saving…"
                                  : epsFridayAttachments.length === 0
                                    ? "No files yet."
                                    : `${epsFridayAttachments.length} file${epsFridayAttachments.length === 1 ? "" : "s"} attached.`}
                              </p>
                            </div>
                            {epsFridayAttachments.length > 0 ? (
                              <ul className="request-list-items">
                                {epsFridayAttachments.map((file) => (
                                  <li
                                    key={file.id}
                                    className="request-list-item"
                                  >
                                    <span className="request-list-label">
                                      {file.name}
                                    </span>
                                    <button
                                      type="button"
                                      className="button secondary"
                                      disabled={
                                        uploadingEpsFridayFiles ||
                                        !isSelectedScheduleActive
                                      }
                                      onClick={() =>
                                        void handleRemoveEpsFridayAttachment(
                                          file.id,
                                        )
                                      }
                                    >
                                      Remove
                                    </button>
                                  </li>
                                ))}
                              </ul>
                            ) : null}
                          </div>
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
                          {showBd15Content ? (
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
                          ) : showBd18Content ? (
                            <>
                              Subject and body are filled from Email Template 7
                              (BD18) — Your Financial Statements Are Ready.
                              Required report files from Attachments are sent
                              with the email
                              {bd15Attachments.length > 0
                                ? ` (${bd15Attachments.length} file${bd15Attachments.length === 1 ? "" : "s"})`
                                : ""}
                              . To is the client POC email; CC is Owner (User
                              Management) and Group ID from Client Profile
                              (same as BD1)
                              {bd15EmailAlreadySent
                                ? ". Sending is disabled because this email was already sent."
                                : "."}
                            </>
                          ) : showBd13Content ? (
                            <>
                              Subject and body are filled from Email Template 5
                              (BD13) — Reminder: Open Questions. Late items
                              mirrored from BD10 fill [Question List]; [Total
                              Amount] sums Amount values on Late rows. Marking
                              Completed here does not change BD9/BD10. Email is
                              available only when at least one Late item remains
                              on BD13. To is the client POC email.
                            </>
                          ) : showBd10Content ? (
                            <>
                              Subject and body are filled from Email Template 6.
                              Outstanding items are all Late rows from BD9 (BD7
                              Status Check and BD8 Status Check). To is the
                              client POC email; CC is Owner (User Management)
                              and Group ID from Client Profile (same as BD1).
                            </>
                          ) : showBd8Content ? (
                            <>
                              Subject and body are filled from Email Template 4
                              (BD8). [Number] is filled Excel/CSV rows minus the
                              header row; [Total Amount] sums the Amount column
                              {questionListAttachment
                                ? ` (${questionListAttachment.name})`
                                : " — add the attachment on Question List Template"}
                              . [Question List Items] comes from questions you
                              add on Question List Template. To is the client
                              POC email. Schedule only from the supervisor
                              account (Admins view only; no Send Now).
                            </>
                          ) : showLastBdSchedule ? (
                            <>
                              Subject and body are filled from Email Template 18
                              (team end-of-month reminder). To is you; BCC is
                              everyone in User Management.
                            </>
                          ) : showBd7Content ? (
                            <>
                              Subject and body are filled from Email Template 3
                              (BD7). Late Action Items come from BD5 for the
                              selected request type. To is the client POC email.
                              Schedule only from the supervisor account (Admins
                              view only; no Send Now).
                            </>
                          ) : showEpsFridayWeek ? (
                            <>
                              Subject and body are filled from Email Template 22
                              (Friday Email). Attachments from the Attachments
                              tab are sent with the email
                              {epsFridayAttachments.length > 0
                                ? ` (${epsFridayAttachments.length} file${epsFridayAttachments.length === 1 ? "" : "s"})`
                                : ""}
                              . Auto-sends at 10:50 am client time. To is the
                              client POC email; CC includes Owner, Group ID, and
                              thomas@ for EPS.
                            </>
                          ) : (
                            <>
                              Subject and body are filled from Email Template{" "}
                              {requestTypeId === "other-statements" ? "2" : "1"}
                              {showBd1TypeContent
                                ? ". To is the client POC email. Schedule only (no Send Now). Request list comes from the Request List Template."
                                : "."}
                            </>
                          )}
                        </p>
                        <div className="client-schedule-email-fields">
                          <label className="field">
                            <span className="field-label">To</span>
                            <input
                              className={`field-input${
                                bd1LikeEmailRecipients
                                  ? " field-input--readonly"
                                  : ""
                              }`}
                              type="text"
                              inputMode="email"
                              value={emailTo}
                              readOnly={bd1LikeEmailRecipients}
                              disabled={bd15EmailAlreadySent}
                              onChange={(event) => {
                                if (bd1LikeEmailRecipients) {
                                  return;
                                }
                                setEmailTo(event.target.value);
                              }}
                              placeholder="email@example.com, email2@example.com"
                            />
                            {bd1LikeEmailRecipients ? (
                              <span className="lookup-hint">
                                {selectedPocEmail
                                  ? `Filled from Client Profile POC${
                                      selectedPointOfContact
                                        ? ` (${selectedPointOfContact})`
                                        : ""
                                    }.`
                                  : "No POC Email on Client Profile — add it under People."}
                              </span>
                            ) : null}
                          </label>
                          <label className="field">
                            <span className="field-label">From</span>
                            <input
                              className="field-input field-input--readonly"
                              type="text"
                              value={
                                emailFrom
                                  ? selectedSupervisorLabel || emailFrom
                                  : "No supervisor assigned"
                              }
                              readOnly
                              tabIndex={-1}
                              aria-readonly="true"
                              title="From is the Supervisor (Primary) mailbox from Resource Management"
                            />
                            {emailSendBlockedReason ? (
                              <span className="lookup-hint">
                                {emailSendBlockedReason}
                              </span>
                            ) : (
                              <span className="lookup-hint">
                                Sends from the supervisor Microsoft account
                                only.
                              </span>
                            )}
                          </label>
                          <label className="field">
                            <span className="field-label">CC</span>
                            <input
                              className={`field-input${
                                bd1LikeEmailRecipients
                                  ? " field-input--readonly"
                                  : ""
                              }`}
                              type="text"
                              inputMode="email"
                              value={emailCc}
                              readOnly={bd1LikeEmailRecipients}
                              disabled={bd15EmailAlreadySent}
                              onChange={(event) => {
                                if (bd1LikeEmailRecipients) {
                                  return;
                                }
                                setEmailCc(event.target.value);
                              }}
                              placeholder="email@example.com, email2@example.com"
                            />
                            {bd1LikeEmailRecipients ? (
                              <span className="lookup-hint">
                                Filled from Owner in User Management
                                {scheduleOwnerCcEmail
                                  ? ` (${scheduleOwnerCcEmail})`
                                  : ""}
                                and Group ID on Client Profile
                                {selectedGroupIdCc
                                  ? ` (${selectedGroupIdCc})`
                                  : " (set Group ID if it should appear here)"}
                                {selectedIsEpsClient
                                  ? `; EPS also includes ${EPS_SCHEDULE_EMAIL_CC}`
                                  : ""}
                                .
                              </span>
                            ) : null}
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
                            {(showBd1TypeContent ||
                              showBd7Content ||
                              showBd8Content ||
                              showBd10Content ||
                              showBd13Content ||
                              showBd15Content ||
                              showBd18Content ||
                              showEpsFridayWeek) &&
                            !showLastBdSchedule ? (
                              <span className="lookup-hint">
                                Saved per client and BD; included on auto-send.
                              </span>
                            ) : null}
                          </label>
                          <label className="field">
                            <span className="field-label">Deadline</span>
                            <input
                              className={`field-input${
                                showBd1TypeContent || showBd7Content
                                  ? " field-input--readonly"
                                  : ""
                              }`}
                              type="date"
                              value={effectiveEmailDeadline}
                              readOnly={showBd1TypeContent || showBd7Content}
                              disabled={bd15EmailAlreadySent}
                              onChange={(event) => {
                                if (showBd1TypeContent || showBd7Content) {
                                  return;
                                }
                                setEmailDeadline(event.target.value);
                              }}
                            />
                            {showBd1TypeContent ? (
                              <span className="lookup-hint">
                                5th business day of the scheduled month (BD5).
                              </span>
                            ) : null}
                            {showBd7Content ? (
                              <span className="lookup-hint">
                                Same as BD7 schedule date.
                              </span>
                            ) : null}
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
                                {showEpsFridayWeek
                                  ? "This Friday at 10:50 am in the client’s time zone"
                                  : "This BD date at 9:00 am in the client’s time zone"}
                              </span>
                            )}
                          </div>
                          <div className="field">
                            <span className="field-label">
                              Scheduled send (Pakistan)
                            </span>
                            <div className="field-input field-input--readonly">
                              {emailSendAtPakistanLabel ||
                                effectiveEmailSendAt ||
                                "—"}
                            </div>
                            <span className="lookup-hint">
                              Same moment as Send at, shown in Pakistan time
                              for the scheduled send
                            </span>
                          </div>
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
                                      {selectedSupervisorEmail
                                        ? `No signature for supervisor (${selectedSupervisorLabel || selectedSupervisorEmail}). They should set it under My Profile while signed in as that account.`
                                        : "Assign a Supervisor in Resource Management to show their signature."}
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
                          {scheduleOnlyEmailDay ? null : (
                            <button
                              type="button"
                              className="button primary"
                              disabled={
                                !canSendAsSupervisor ||
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
                          )}
                          {scheduleOnlyEmailDay ? null : (
                            <button
                              type="button"
                              className="button secondary-solid"
                              disabled={
                                !canSendAsSupervisor ||
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
                          )}
                          {scheduleOnlyEmailDay ? (
                            <span className="lookup-hint">
                              {showEpsFridayWeek
                                ? `${selectedSchedule.label} auto-sends at 10:50 am client time on that Friday from the supervisor mailbox (server cron). One EPS Fridays star enables all weeks. Schedule is disabled.`
                                : `${selectedScheduleId.toUpperCase()} auto-sends at 9:00 am client time on its BD date from the supervisor mailbox (server cron). Schedule is disabled for everyone.`}
                            </span>
                          ) : null}
                        </div>
                        {scheduleOnlyEmailDay &&
                        bdAutoEmailStatus.attempts.length > 0 ? (
                          <div className="bd-auto-email-attempts">
                            <p className="lookup-hint" style={{ marginBottom: 6 }}>
                              Auto-send checks (after send time)
                            </p>
                            <ul className="lookup-hint" style={{ margin: 0, paddingLeft: "1.2em" }}>
                              {bdAutoEmailStatus.attempts.map((attempt) => {
                                const when =
                                  formatActionResolvedAt(attempt.at) ||
                                  attempt.at ||
                                  "—";
                                if (attempt.outcome === "sent") {
                                  return (
                                    <li key={`${attempt.at}-sent`}>
                                      Check at {when}
                                      {attempt.fromEmail
                                        ? ` from ${attempt.fromEmail}`
                                        : ""}
                                      : sent.
                                    </li>
                                  );
                                }
                                return (
                                  <li key={`${attempt.at}-error`}>
                                    Check at {when}: auto-send failed
                                    {attempt.message
                                      ? ` — ${attempt.message}`
                                      : "."}
                                  </li>
                                );
                              })}
                            </ul>
                            {bdAutoEmailStatus.emailSent ? (
                              <p className="form-message success">
                                Last sent at{" "}
                                {formatActionResolvedAt(
                                  bdAutoEmailStatus.lastSentAt,
                                ) ||
                                  bdAutoEmailStatus.lastSentAt ||
                                  "—"}
                                {bdAutoEmailStatus.fromEmail
                                  ? ` from ${bdAutoEmailStatus.fromEmail}`
                                  : ""}
                                .
                              </p>
                            ) : null}
                          </div>
                        ) : null}
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

                      const requestBodyTitle = showBd15Content
                        ? "Books Closed."
                        : showBd18Content
                          ? "Your Financial Statements Are Ready"
                          : showBd13Content
                            ? "Reminder: Open Questions"
                            : showBd10Content
                              ? "Reminder: Final Reminder for Missing Documents"
                              : showBd9Content
                                ? "Manual Check: Uncategorized Transactions: Questions for the Client"
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
