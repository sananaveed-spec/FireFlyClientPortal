import { isGraphAppOnlyConfigured } from "@/lib/azure-graph-app";
import {
  BD_AUTO_SEND_SCHEDULE_IDS,
  getBdAutoEmailStatus,
  isBdAutoSendScheduleId,
  markBdAutoEmailError,
  markBdAutoEmailSent,
  normalizeBdAutoEmailStatusStore,
  type BdAutoEmailStatusStore,
  type BdAutoSendScheduleId,
} from "@/lib/bd-auto-email-status";
import {
  getBd8EmailAttachment,
  normalizeBd8EmailAttachmentsStore,
} from "@/lib/bd8-email-attachments";
import {
  getBd15EmailAttachments,
  normalizeBd15EmailAttachmentsStore,
} from "@/lib/bd15-email-attachments";
import {
  getBd15EmailStatus,
  markBd15EmailSent,
  normalizeBd15EmailStatusStore,
} from "@/lib/bd15-email-status";
import {
  getBdEmailBcc,
  normalizeBdEmailBccStore,
} from "@/lib/bd-email-bcc";
import {
  getScheduleBusinessDate,
  getUpcomingScheduleBusinessDate,
  formatBusinessDayLabel,
} from "@/lib/business-days";
import {
  getEpsFridayAttachments,
  normalizeEpsFridayAttachmentsStore,
} from "@/lib/eps-friday-attachments";
import {
  EPS_FRIDAYS_TOGGLE_ID,
  getEpsFridayWeekItemsForMonth,
  isEpsFridayWeekScheduleId,
  isEpsFridaysActive,
} from "@/lib/eps-friday-weeks";
import { findProfileForResourceClient } from "@/lib/client-management";
import {
  bd1TemplateToRequestTypeId,
  normalizeBd1Template,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import {
  BD13_LATE_SOURCE_PREFIX,
  BD9_EXCEL_SOURCE_PREFIX,
  currentMonthKey,
  getClientRequestList,
  normalizeClientRequestListsStore,
  requestItemsForMonth,
  type ActionWorkflowItem,
  type ClientRequestListsStore,
} from "@/lib/client-request-lists";
import {
  getActiveScheduleIds,
  normalizeClientScheduleSettingsStore,
} from "@/lib/client-schedule-settings";
import {
  readBd8EmailAttachmentsFile,
  readBd15EmailAttachmentsFile,
  readBd15EmailStatusFile,
  readBdAutoEmailStatusFile,
  readBdEmailBccFile,
  readClientProfiles,
  readClientRequestListsFile,
  readClientScheduleSettingsFile,
  readEmailSignatureForUser,
  readEmailTemplates,
  readEpsFridayAttachmentsFile,
  readHolidays,
  readPortalUsersFile,
  readResourceClients,
  writeBd15EmailStatusFile,
  writeBdAutoEmailStatusFile,
} from "@/lib/data-store";
import {
  applyPeriodPlaceholder,
  emailTemplateLabelNumber,
  type EmailTemplate,
  type EmailTemplatesStore,
} from "@/lib/email-templates";
import {
  applyPortalLinkPlaceholders,
  resolvePortalHref,
} from "@/lib/email-portal-link";
import {
  createEmptyEmailSignature,
  stripLegacySignatureClosing,
} from "@/lib/email-signatures";
import type { HolidaysStore } from "@/lib/holidays";
import { parsePortalUsers, type PortalUser } from "@/lib/portal-users";
import { sendGraphAppMail } from "@/lib/send-graph-app-mail";
import type { GraphMailAttachment } from "@/lib/send-graph-mail";
import { resolveStaffEmail } from "@/lib/staff-assign";
import {
  calendarYmdFromLocalDate,
  getCalendarYmdInTimeZone,
  getClientEpsFridaySendSchedule,
  getClientNineAmSchedule,
  resolveClientTimeZone,
} from "@/lib/time-zones";

const SCHEDULE_EMAIL_OWNER_CC_FALLBACK = "ayesha@fireflyaccountancy.com";
const EPS_SCHEDULE_EMAIL_CC = "thomas@epsfresno.com";

/** How long after sendAt a cron may still send (covers EST→PST same day). */
const SEND_WINDOW_MS = 18 * 60 * 60 * 1000;

export type BdAutoEmailRunResult = {
  ok: boolean;
  checkedAt: string;
  sent: Array<{
    clientId: string;
    clientName: string;
    scheduleId: BdAutoSendScheduleId;
    from: string;
    to: string;
    lastSentAt: string;
  }>;
  skipped: Array<{
    clientId: string;
    clientName: string;
    scheduleId: BdAutoSendScheduleId;
    reason: string;
  }>;
  errors: Array<{
    clientId: string;
    clientName: string;
    scheduleId: BdAutoSendScheduleId;
    error: string;
  }>;
};

function resolveOwnerEmail(users: PortalUser[]): string {
  const owner = users.find(
    (user) => user.position?.trim().toLowerCase() === "owner",
  );
  return owner?.email?.trim() || SCHEDULE_EMAIL_OWNER_CC_FALLBACK;
}

function isEpsClient(profile: ClientProfileEntry): boolean {
  const legal = profile.legalName.trim().toLowerCase();
  const operating = profile.operatingName.trim().toLowerCase();
  // Match "Electrical Power System" / "Systems" / short "EPS".
  if (
    /electrical power systems?\b/.test(legal) ||
    /electrical power systems?\b/.test(operating) ||
    operating === "eps" ||
    legal === "eps"
  ) {
    return true;
  }
  return /\/sites\/eps(\/|$)/i.test(profile.portalLink);
}

function contactFirstName(profile: ClientProfileEntry): string {
  const poc = profile.pointOfContact.trim();
  if (!poc) {
    return "";
  }
  return poc.split(/\s+/)[0] || poc;
}

function findTemplateById(
  store: EmailTemplatesStore,
  templateId: string,
): EmailTemplate | null {
  return store.templates.find((template) => template.id === templateId) ?? null;
}

function findTemplateByLabelNumber(
  store: EmailTemplatesStore,
  labelNumber: number,
): EmailTemplate | null {
  return (
    store.templates.find(
      (template) => emailTemplateLabelNumber(template) === labelNumber,
    ) ?? null
  );
}

function bd1TemplateForProfile(
  store: EmailTemplatesStore,
  profile: ClientProfileEntry,
): EmailTemplate | null {
  const requestType = bd1TemplateToRequestTypeId(
    normalizeBd1Template(profile.bd1Template),
  );
  const id =
    requestType === "other-statements" ? "template-2" : "template-1";
  return findTemplateById(store, id);
}

function templateForSchedule(
  store: EmailTemplatesStore,
  scheduleId: BdAutoSendScheduleId,
  profile: ClientProfileEntry,
): EmailTemplate | null {
  if (isEpsFridayWeekScheduleId(String(scheduleId))) {
    // Prefer live Lookup Data Template 22 (label), then seed id.
    return (
      findTemplateByLabelNumber(store, 22) ??
      findTemplateById(store, "template-22")
    );
  }
  if (scheduleId === "bd1") {
    return bd1TemplateForProfile(store, profile);
  }
  if (scheduleId === "bd7") {
    return findTemplateByLabelNumber(store, 3);
  }
  if (scheduleId === "bd8") {
    return findTemplateByLabelNumber(store, 4);
  }
  if (scheduleId === "bd13") {
    return findTemplateByLabelNumber(store, 5);
  }
  if (scheduleId === "bd10") {
    return findTemplateByLabelNumber(store, 6);
  }
  if (scheduleId === "bd18") {
    return findTemplateByLabelNumber(store, 7);
  }
  if (scheduleId === "bd15") {
    return findTemplateById(store, "template-11") ??
      findTemplateByLabelNumber(store, 11);
  }
  return null;
}

function requestListScheduleIdForBd1(profile: ClientProfileEntry): string {
  const requestType = bd1TemplateToRequestTypeId(
    normalizeBd1Template(profile.bd1Template),
  );
  return requestType === "other-statements"
    ? "bd1-other-statements"
    : "bd1";
}

function getBd5LateItems(
  store: ClientRequestListsStore,
  clientId: string,
  monthKey: string,
): ActionWorkflowItem[] {
  const lateById = new Map<string, ActionWorkflowItem>();
  for (const scheduleId of ["bd5", "bd5-other-statements"]) {
    const list = getClientRequestList(store, clientId, scheduleId);
    for (const item of list.actionItems) {
      if (item.monthKey === monthKey && item.status === "late") {
        lateById.set(item.id, item);
      }
    }
  }
  return Array.from(lateById.values());
}

/** BD10: Late from BD5 (BD7 Status Check) + BD9 Excel Status Check. */
function getBd10LateItems(
  store: ClientRequestListsStore,
  clientId: string,
  monthKey: string,
): ActionWorkflowItem[] {
  const lateById = new Map<string, ActionWorkflowItem>();
  for (const item of getBd5LateItems(store, clientId, monthKey)) {
    lateById.set(item.id, item);
  }
  const bd9 = getClientRequestList(store, clientId, "bd9");
  for (const item of bd9.actionItems) {
    if (
      item.monthKey === monthKey &&
      item.status === "late" &&
      item.sourceRequestId.startsWith(BD9_EXCEL_SOURCE_PREFIX)
    ) {
      lateById.set(item.id, item);
    }
  }
  return Array.from(lateById.values());
}

/** BD13: local late mirrors only. */
function getBd13LateItems(
  store: ClientRequestListsStore,
  clientId: string,
  monthKey: string,
): ActionWorkflowItem[] {
  const list = getClientRequestList(store, clientId, "bd13");
  return list.actionItems.filter(
    (item) =>
      item.monthKey === monthKey &&
      item.status === "late" &&
      item.sourceRequestId.startsWith(BD13_LATE_SOURCE_PREFIX),
  );
}

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
  if (reqKey === "balance sheet") {
    return fileKey.includes("balance sheet");
  }
  if (reqKey.includes("profit") && reqKey.includes("loss")) {
    return fileKey.includes("profit") && fileKey.includes("loss");
  }
  const requiredWords = reqKey.split(" ").filter(Boolean);
  return (
    requiredWords.length > 0 &&
    requiredWords.every((word) => fileKey.includes(word))
  );
}

function getMissingBd15Reports(
  attachments: ReadonlyArray<{ name: string }>,
): string[] {
  const required = ["Profit & Loss Account", "Balance Sheet"] as const;
  return required.filter(
    (label) =>
      !attachments.some((file) => bd15FileMatchesReport(file.name, label)),
  );
}

function toBulletList(labels: string[]): string {
  return labels
    .map((label) => label.trim())
    .filter(Boolean)
    .map((label) => `• ${label}`)
    .join("\n");
}

function formatDueDateLabel(date: Date | null): string {
  if (!date) {
    return "[Due Date]";
  }
  return formatBusinessDayLabel(date);
}

/** Short date for Template 22 `[date]` (e.g. Nov 7, 2026). */
function formatEpsFridayDateLabel(date: Date | null): string {
  if (!date) {
    return "[date]";
  }
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  });
}

function publicPortalOrigin(): string {
  const raw =
    process.env.IGNITION_PUBLIC_BASE_URL?.trim() ||
    process.env.NEXT_PUBLIC_PORTAL_BASE_URL?.trim() ||
    "";
  if (!raw) {
    return "";
  }
  try {
    return new URL(raw).origin;
  } catch {
    return raw.replace(/\/+$/, "");
  }
}

function authorizeCron(request: Request): boolean {
  const secret = process.env.BD_AUTO_EMAIL_CRON_SECRET?.trim();
  if (!secret) {
    return false;
  }
  const header = request.headers.get("authorization")?.trim() ?? "";
  const bearer = header.toLowerCase().startsWith("bearer ")
    ? header.slice(7).trim()
    : "";
  const alt = request.headers.get("x-cron-secret")?.trim() ?? "";
  return bearer === secret || alt === secret;
}

export function isBdAutoEmailCronAuthorized(request: Request): boolean {
  return authorizeCron(request);
}

function buildEmailContent(input: {
  scheduleId: BdAutoSendScheduleId;
  profile: ClientProfileEntry;
  template: EmailTemplate;
  requestLists: ClientRequestListsStore;
  holidays: HolidaysStore;
  now: Date;
  clientId: string;
}): { subject: string; body: string; portalHref: string | null } {
  const { scheduleId, profile, template, requestLists, holidays, now, clientId } =
    input;
  const monthKey = currentMonthKey(now);
  const clientName =
    profile.operatingName.replace(/[.:]+$/g, "").trim() ||
    profile.operatingName.trim() ||
    profile.legalName.trim() ||
    "[Client Name]";
  const contactName =
    contactFirstName(profile) ||
    profile.pointOfContact.trim() ||
    "[Contact First Name]";
  const portalLink = profile.portalLink.trim() || "[Portal Link]";

  const bd1Date = getUpcomingScheduleBusinessDate(
    "bd1",
    holidays,
    profile.country,
    profile.state,
    now,
  );
  // BD1 deadline = 5th business day of the same month as BD1.
  const dueDate =
    scheduleId === "bd1" && bd1Date
      ? getScheduleBusinessDate(
          "bd5",
          holidays,
          profile.country,
          profile.state,
          bd1Date,
        )
      : getUpcomingScheduleBusinessDate(
          scheduleId,
          holidays,
          profile.country,
          profile.state,
          now,
        );
  const dueDateLabel = formatDueDateLabel(dueDate);
  const bd1DateLabel = formatDueDateLabel(bd1Date);
  const fridayDateLabel = formatEpsFridayDateLabel(dueDate);
  const periodDate = now;
  const isEpsFriday = isEpsFridayWeekScheduleId(String(scheduleId));

  let emailLabels: string[] = [];
  if (scheduleId === "bd7") {
    emailLabels = getBd5LateItems(requestLists, clientId, monthKey).map(
      (item) => item.label,
    );
  } else if (scheduleId === "bd10") {
    emailLabels = getBd10LateItems(requestLists, clientId, monthKey).map(
      (item) => item.label,
    );
  } else if (scheduleId === "bd13") {
    emailLabels = getBd13LateItems(requestLists, clientId, monthKey).map(
      (item) => item.label,
    );
  } else if (scheduleId === "bd8") {
    const list = getClientRequestList(requestLists, clientId, "bd8");
    const period = requestItemsForMonth(list.requestItems, monthKey);
    emailLabels = (period.length > 0 ? period : list.requestItems).map(
      (item) => item.label,
    );
  } else if (
    scheduleId === "bd15" ||
    scheduleId === "bd18" ||
    isEpsFriday
  ) {
    emailLabels = [];
  } else {
    const listId = requestListScheduleIdForBd1(profile);
    const list = getClientRequestList(requestLists, clientId, listId);
    const period = requestItemsForMonth(list.requestItems, monthKey);
    emailLabels = (period.length > 0 ? period : list.requestItems).map(
      (item) => item.label,
    );
  }

  const usesLateList =
    scheduleId === "bd7" || scheduleId === "bd10" || scheduleId === "bd13";
  const actionListText =
    emailLabels.length > 0
      ? toBulletList(emailLabels)
      : usesLateList
        ? "• [Late Action Items]"
        : "• [Request List Items]";
  const questionListText =
    emailLabels.length > 0 ? toBulletList(emailLabels) : "• [Question List]";
  const questionCount =
    emailLabels.length > 0 ? String(emailLabels.length) : "[Number]";

  function fillPlaceholders(text: string) {
    return applyPeriodPlaceholder(
      applyPortalLinkPlaceholders(
        text
          .replace(/\[Client Name\]/gi, clientName)
          .replace(/\[Client\]/gi, clientName)
          .replace(/\[Contact First Name\]/gi, contactName)
          .replace(/\[Due Date\]/gi, dueDateLabel)
          .replace(/\[Deadline\]/gi, dueDateLabel)
          .replace(/\[BD1 Date\]/gi, bd1DateLabel)
          .replace(/\[date\]/gi, fridayDateLabel)
          .replace(/\[Number\]/gi, questionCount)
          .replace(/\[Total Amount\]/gi, "[Total Amount]")
          .replace(/(?:•\t*\s*)?\[Question List Items\]/gi, questionListText)
          .replace(/(?:•\s*)?\[Question List\]/gi, questionListText)
          .replace(/(?:•\t*\s*)?\[Request List Items\]/gi, actionListText)
          .replace(/(?:•\t*\s*)?\[Late Action Items\]/gi, actionListText)
          .replace(/(?:•\t*\s*)?\[Outstanding items\]/gi, actionListText)
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

  let subject = fillPlaceholders(template.subject?.trim() || "");
  let body = fillPlaceholders(
    stripLegacySignatureClosing(template.body ?? "").trimEnd(),
  );

  if (
    scheduleId === "bd1" &&
    emailLabels.length > 0 &&
    !emailLabels.every((label) => body.includes(label))
  ) {
    // Insert list once after the "accounts below:" line when placeholders
    // were missing — keep a single blank line, not a large gap.
    const injected = body.replace(
      /(need the statements(?: for the accounts)? below:)\s*/i,
      `$1\n${actionListText}\n`,
    );
    body =
      injected !== body
        ? injected
        : `${body.trimEnd()}\n${actionListText}`;
  }

  // Collapse 3+ newlines to at most one blank line (tightens list → portal gap).
  body = body.replace(/\n{3,}/g, "\n\n").trimEnd();

  const portalHref = resolvePortalHref(portalLink, periodDate) || null;
  return { subject, body, portalHref };
}

/**
 * Run auto-send for Active email BDs whose calendar date is today (client TZ)
 * and whose 9:00 AM send time has been reached. Other BDs are skipped quietly.
 */
export async function runBdAutoEmails(options?: {
  force?: boolean;
  now?: Date;
  scheduleIds?: BdAutoSendScheduleId[];
}): Promise<BdAutoEmailRunResult> {
  const now = options?.now ?? new Date();
  const force = Boolean(options?.force);
  const epsWeekIds = getEpsFridayWeekItemsForMonth(now).map((item) => item.id);
  const scheduleIds = options?.scheduleIds?.length
    ? options.scheduleIds
    : [...BD_AUTO_SEND_SCHEDULE_IDS, ...epsWeekIds];
  const checkedAt = now.toISOString();
  const knownActiveIds = [
    ...BD_AUTO_SEND_SCHEDULE_IDS,
    EPS_FRIDAYS_TOGGLE_ID,
    ...epsWeekIds,
  ];
  const result: BdAutoEmailRunResult = {
    ok: true,
    checkedAt,
    sent: [],
    skipped: [],
    errors: [],
  };

  if (!isGraphAppOnlyConfigured()) {
    result.ok = false;
    result.errors.push({
      clientId: "",
      clientName: "",
      scheduleId: "bd1",
      error:
        "AZURE_CLIENT_SECRET is not set (required for app-only Mail.Send).",
    });
    return result;
  }

  const [
    profiles,
    resources,
    holidays,
    templates,
    usersRaw,
    requestListsRaw,
    scheduleSettingsRaw,
    statusRaw,
    bd8AttachmentsRaw,
    bd15AttachmentsRaw,
    bd15StatusRaw,
    bdEmailBccRaw,
    epsFridayAttachmentsRaw,
  ] = await Promise.all([
    readClientProfiles(),
    readResourceClients(),
    readHolidays(),
    readEmailTemplates(),
    readPortalUsersFile(),
    readClientRequestListsFile(),
    readClientScheduleSettingsFile(),
    readBdAutoEmailStatusFile(),
    readBd8EmailAttachmentsFile(),
    readBd15EmailAttachmentsFile(),
    readBd15EmailStatusFile(),
    readBdEmailBccFile(),
    readEpsFridayAttachmentsFile(),
  ]);

  const users = parsePortalUsers(usersRaw);
  const requestLists = normalizeClientRequestListsStore(requestListsRaw);
  const scheduleSettings =
    normalizeClientScheduleSettingsStore(scheduleSettingsRaw);
  const monthKey = currentMonthKey(now);
  let statusStore = normalizeBdAutoEmailStatusStore(statusRaw, monthKey);
  const bd8Attachments = normalizeBd8EmailAttachmentsStore(
    bd8AttachmentsRaw,
    monthKey,
  );
  const bd15Attachments = normalizeBd15EmailAttachmentsStore(
    bd15AttachmentsRaw,
    monthKey,
  );
  let bd15StatusStore = normalizeBd15EmailStatusStore(
    bd15StatusRaw,
    monthKey,
  );
  let bd15StatusDirty = false;
  const bdEmailBccStore = normalizeBdEmailBccStore(bdEmailBccRaw);
  const epsFridayAttachmentsStore = normalizeEpsFridayAttachmentsStore(
    epsFridayAttachmentsRaw,
    monthKey,
  );
  const ownerCc = resolveOwnerEmail(users);
  const publicOrigin = publicPortalOrigin();

  for (const resource of resources) {
    const clientName = resource.clientName.trim() || resource.id;
    const profile = findProfileForResourceClient(resource, profiles);

    if (!profile) {
      continue;
    }

    const clientIana = resolveClientTimeZone(profile.timeZone);
    if (!clientIana) {
      continue;
    }
    const todayYmd = getCalendarYmdInTimeZone(now, clientIana);

    const activeIds = getActiveScheduleIds(
      scheduleSettings,
      resource.id,
      knownActiveIds,
    );
    const epsClient = isEpsClient(profile);

    for (const scheduleId of scheduleIds) {
      const isEpsFriday = isEpsFridayWeekScheduleId(scheduleId);
      if (isEpsFriday && !epsClient) {
        continue;
      }

      // Only evaluate schedules whose this-month date equals today in client TZ.
      const bdDate = getScheduleBusinessDate(
        scheduleId,
        holidays,
        profile.country,
        profile.state,
        now,
      );
      if (!bdDate || calendarYmdFromLocalDate(bdDate) !== todayYmd) {
        continue;
      }

      const isActive = isEpsFriday
        ? isEpsFridaysActive(activeIds)
        : activeIds.includes(scheduleId);
      if (!isActive) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: isEpsFriday
            ? "EPS Fridays is Inactive (star not filled)."
            : `${scheduleId.toUpperCase()} is Inactive (star not filled).`,
        });
        continue;
      }

      const existing = getBdAutoEmailStatus(
        statusStore,
        resource.id,
        scheduleId,
        monthKey,
      );
      if (!force && existing?.emailSent) {
        // Already sent — no attempt history / UI noise.
        continue;
      }

      if (
        (scheduleId === "bd15" || scheduleId === "bd18") &&
        !force &&
        getBd15EmailStatus(bd15StatusStore, resource.id, monthKey).emailSent
      ) {
        continue;
      }

      const supervisorEmail = resolveStaffEmail(resource.supervisor, users);
      if (!supervisorEmail) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: "No supervisor email on Resource Allocation.",
        });
        continue;
      }

      const to = profile.pocEmail.trim();
      if (!to) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: "No POC Email on Client Profile.",
        });
        continue;
      }

      const sendSchedule = isEpsFriday
        ? getClientEpsFridaySendSchedule(bdDate, profile.timeZone)
        : getClientNineAmSchedule(bdDate, profile.timeZone);
      if (!sendSchedule) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: "Client time zone missing or invalid.",
        });
        continue;
      }

      const sendAt = sendSchedule.instant;
      if (now.getTime() < sendAt.getTime()) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: `Send time not reached yet (${sendAt.toISOString()}).`,
        });
        continue;
      }
      if (now.getTime() > sendAt.getTime() + SEND_WINDOW_MS) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: `Outside send window (sendAt ${sendAt.toISOString()}).`,
        });
        continue;
      }

      if (scheduleId === "bd7") {
        const late = getBd5LateItems(requestLists, resource.id, monthKey);
        if (late.length === 0) {
          result.skipped.push({
            clientId: resource.id,
            clientName,
            scheduleId,
            reason: "No BD5 Late items for this month.",
          });
          continue;
        }
      }

      if (scheduleId === "bd10") {
        const late = getBd10LateItems(requestLists, resource.id, monthKey);
        if (late.length === 0) {
          result.skipped.push({
            clientId: resource.id,
            clientName,
            scheduleId,
            reason: "No late BD10 question list items for this month.",
          });
          continue;
        }
      }

      if (scheduleId === "bd13") {
        const late = getBd13LateItems(requestLists, resource.id, monthKey);
        if (late.length === 0) {
          result.skipped.push({
            clientId: resource.id,
            clientName,
            scheduleId,
            reason: "No late BD13 question list items for this month.",
          });
          continue;
        }
      }

      const mailAttachments: GraphMailAttachment[] = [];
      if (scheduleId === "bd8") {
        const file = getBd8EmailAttachment(
          bd8Attachments,
          resource.id,
          monthKey,
          "bd8",
        );
        if (!file) {
          result.skipped.push({
            clientId: resource.id,
            clientName,
            scheduleId,
            reason: "BD8 Excel/CSV attachment missing.",
          });
          continue;
        }
        mailAttachments.push({
          name: file.name,
          contentType: file.contentType,
          contentBytes: file.contentBytes,
        });
      }

      if (isEpsFriday) {
        for (const file of getEpsFridayAttachments(
          epsFridayAttachmentsStore,
          resource.id,
          scheduleId,
          monthKey,
        )) {
          mailAttachments.push({
            name: file.name,
            contentType: file.contentType,
            contentBytes: file.contentBytes,
          });
        }
      }

      if (scheduleId === "bd15" || scheduleId === "bd18") {
        const files = getBd15EmailAttachments(
          bd15Attachments,
          resource.id,
          monthKey,
        );
        const missing = getMissingBd15Reports(files);
        if (missing.length > 0) {
          result.skipped.push({
            clientId: resource.id,
            clientName,
            scheduleId,
            reason: `Missing BD15 reports: ${missing.join(", ")}.`,
          });
          continue;
        }
        for (const file of files) {
          mailAttachments.push({
            name: file.name,
            contentType: file.contentType,
            contentBytes: file.contentBytes,
          });
        }
      }

      const template = templateForSchedule(templates, scheduleId, profile);
      if (!template) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: "Email template not found.",
        });
        continue;
      }

      const content = buildEmailContent({
        scheduleId,
        profile,
        template,
        requestLists,
        holidays,
        now,
        clientId: resource.id,
      });
      if (!content.subject.trim()) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: "Empty subject after template fill.",
        });
        continue;
      }

      const fromEmail = supervisorEmail.toLowerCase();
      const toEmail = to.toLowerCase();
      const ccParts = [
        ownerCc,
        profile.groupId.trim(),
        isEpsClient(profile) ? EPS_SCHEDULE_EMAIL_CC : "",
      ]
        .map((email) => email.trim().toLowerCase())
        .filter(
          (email) => Boolean(email) && email !== toEmail && email !== fromEmail,
        );
      const cc = [...new Set(ccParts)].join(", ");
      const savedBcc = getBdEmailBcc(
        bdEmailBccStore,
        resource.id,
        scheduleId,
      );
      const bcc = [
        ...new Set(
          savedBcc
            .split(/[;,]+/)
            .map((email) => email.trim().toLowerCase())
            .filter(
              (email) =>
                Boolean(email) &&
                email !== toEmail &&
                email !== fromEmail &&
                !ccParts.includes(email),
            ),
        ),
      ].join(", ");

      let signature = await readEmailSignatureForUser(supervisorEmail);
      const supervisorLabel = resource.supervisor.trim();
      const supervisorDisplayName =
        supervisorLabel.includes("@") ? "" : supervisorLabel;
      if (!signature.email.trim()) {
        signature = {
          ...signature,
          email: supervisorEmail,
          name: signature.name.trim() || supervisorDisplayName,
        };
      }
      if (
        !signature.name.trim() &&
        !signature.position.trim() &&
        !signature.phone.trim()
      ) {
        signature = {
          ...createEmptyEmailSignature(),
          name: supervisorDisplayName,
          email: supervisorEmail,
        };
      }
      if (
        signature.name.trim().toLowerCase() ===
        signature.email.trim().toLowerCase()
      ) {
        signature = { ...signature, name: "" };
      }

      try {
        await sendGraphAppMail({
          from: supervisorEmail,
          to,
          cc,
          bcc,
          subject: content.subject,
          bodyText: content.body,
          portalHref: content.portalHref,
          signature,
          publicOrigin,
          attachments: mailAttachments,
        });
        const lastSentAt = now.toISOString();
        statusStore = markBdAutoEmailSent(statusStore, {
          clientId: resource.id,
          scheduleId,
          monthKey,
          lastSentAt,
          fromEmail: supervisorEmail,
          source: "auto-cron",
        });
        if (scheduleId === "bd15" || scheduleId === "bd18") {
          bd15StatusStore = markBd15EmailSent(
            bd15StatusStore,
            resource.id,
            monthKey,
            lastSentAt,
          );
          bd15StatusDirty = true;
        }
        result.sent.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          from: supervisorEmail,
          to,
          lastSentAt,
        });
      } catch (sendError) {
        const message =
          sendError instanceof Error
            ? sendError.message
            : "Failed to send BD auto email.";
        statusStore = markBdAutoEmailError(statusStore, {
          clientId: resource.id,
          scheduleId,
          monthKey,
          lastError: message,
          fromEmail: supervisorEmail,
          attemptedAt: now.toISOString(),
        });
        result.errors.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          error: message,
        });
        result.ok = false;
      }
    }
  }

  await writeBdAutoEmailStatusFile(statusStore);
  if (bd15StatusDirty) {
    await writeBd15EmailStatusFile(bd15StatusStore);
  }
  return result;
}

export async function readNormalizedBdAutoEmailStatus(
  monthKey?: string,
): Promise<BdAutoEmailStatusStore> {
  const raw = await readBdAutoEmailStatusFile();
  return normalizeBdAutoEmailStatusStore(
    raw,
    monthKey?.trim() || currentMonthKey(),
  );
}

export { isBdAutoSendScheduleId };
