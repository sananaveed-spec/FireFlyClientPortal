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
import { getUpcomingScheduleBusinessDate } from "@/lib/business-days";
import { findProfileForResourceClient } from "@/lib/client-management";
import {
  bd1TemplateToRequestTypeId,
  normalizeBd1Template,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import {
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
  readBdAutoEmailStatusFile,
  readClientProfiles,
  readClientRequestListsFile,
  readClientScheduleSettingsFile,
  readEmailSignatureForUser,
  readEmailTemplates,
  readHolidays,
  readPortalUsersFile,
  readResourceClients,
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
import { formatBusinessDayLabel } from "@/lib/business-days";
import type { HolidaysStore } from "@/lib/holidays";
import { parsePortalUsers, type PortalUser } from "@/lib/portal-users";
import { sendGraphAppMail } from "@/lib/send-graph-app-mail";
import { resolveStaffEmail } from "@/lib/staff-assign";
import { getClientNineAmSchedule } from "@/lib/time-zones";

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
  if (
    legal.includes("electrical power systems") ||
    operating.includes("electrical power systems") ||
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
  if (scheduleId === "bd1") {
    return bd1TemplateForProfile(store, profile);
  }
  if (scheduleId === "bd7") {
    return findTemplateByLabelNumber(store, 3);
  }
  return findTemplateByLabelNumber(store, 4);
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

  const dueDate = getUpcomingScheduleBusinessDate(
    scheduleId === "bd1" ? "bd4" : scheduleId,
    holidays,
    profile.country,
    profile.state,
    now,
  );
  const dueDateLabel = formatDueDateLabel(dueDate);
  const bd1Date = getUpcomingScheduleBusinessDate(
    "bd1",
    holidays,
    profile.country,
    profile.state,
    now,
  );
  const bd1DateLabel = formatDueDateLabel(bd1Date);
  const periodDate = now;

  let emailLabels: string[] = [];
  if (scheduleId === "bd7") {
    emailLabels = getBd5LateItems(requestLists, clientId, monthKey).map(
      (item) => item.label,
    );
  } else if (scheduleId === "bd8") {
    const list = getClientRequestList(requestLists, clientId, "bd8");
    const period = requestItemsForMonth(list.requestItems, monthKey);
    emailLabels = (period.length > 0 ? period : list.requestItems).map(
      (item) => item.label,
    );
  } else {
    const listId = requestListScheduleIdForBd1(profile);
    const list = getClientRequestList(requestLists, clientId, listId);
    const period = requestItemsForMonth(list.requestItems, monthKey);
    emailLabels = (period.length > 0 ? period : list.requestItems).map(
      (item) => item.label,
    );
  }

  const actionListText =
    emailLabels.length > 0
      ? toBulletList(emailLabels)
      : scheduleId === "bd7"
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
    body = `${body.trimEnd()}\n\n${actionListText}`;
  }

  const portalHref = resolvePortalHref(portalLink, periodDate) || null;
  return { subject, body, portalHref };
}

/**
 * Run auto-send for Active BD1 / BD7 / BD8 clients whose send time has been
 * reached and who are not already marked sent for the month.
 */
export async function runBdAutoEmails(options?: {
  force?: boolean;
  now?: Date;
  scheduleIds?: BdAutoSendScheduleId[];
}): Promise<BdAutoEmailRunResult> {
  const now = options?.now ?? new Date();
  const force = Boolean(options?.force);
  const scheduleIds = options?.scheduleIds?.length
    ? options.scheduleIds
    : [...BD_AUTO_SEND_SCHEDULE_IDS];
  const checkedAt = now.toISOString();
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
  ] = await Promise.all([
    readClientProfiles(),
    readResourceClients(),
    readHolidays(),
    readEmailTemplates(),
    readPortalUsersFile(),
    readClientRequestListsFile(),
    readClientScheduleSettingsFile(),
    readBdAutoEmailStatusFile(),
  ]);

  const users = parsePortalUsers(usersRaw);
  const requestLists = normalizeClientRequestListsStore(requestListsRaw);
  const scheduleSettings =
    normalizeClientScheduleSettingsStore(scheduleSettingsRaw);
  let statusStore = normalizeBdAutoEmailStatusStore(
    statusRaw,
    currentMonthKey(now),
  );
  const ownerCc = resolveOwnerEmail(users);
  const publicOrigin = publicPortalOrigin();

  for (const resource of resources) {
    const clientName = resource.clientName.trim() || resource.id;
    const profile = findProfileForResourceClient(resource, profiles);

    if (!profile) {
      for (const scheduleId of scheduleIds) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: "No matching Client Profile.",
        });
      }
      continue;
    }

    const activeIds = getActiveScheduleIds(
      scheduleSettings,
      resource.id,
      BD_AUTO_SEND_SCHEDULE_IDS as unknown as string[],
    );

    for (const scheduleId of scheduleIds) {
      if (!activeIds.includes(scheduleId)) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: `${scheduleId.toUpperCase()} is Inactive (star not filled).`,
        });
        continue;
      }

      const monthKey = currentMonthKey(now);
      const existing = getBdAutoEmailStatus(
        statusStore,
        resource.id,
        scheduleId,
        monthKey,
      );
      if (!force && existing?.emailSent) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: `Already sent at ${existing.lastSentAt || "unknown"}.`,
        });
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

      const bdDate = getUpcomingScheduleBusinessDate(
        scheduleId,
        holidays,
        profile.country,
        profile.state,
        now,
      );
      const sendSchedule = getClientNineAmSchedule(bdDate, profile.timeZone);
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
      // ?force=1 bypasses the clock window so you can test immediately.
      if (!force && now.getTime() < sendAt.getTime()) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId,
          reason: `Send time not reached yet (${sendAt.toISOString()}).`,
        });
        continue;
      }
      if (!force && now.getTime() > sendAt.getTime() + SEND_WINDOW_MS) {
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

      let signature = await readEmailSignatureForUser(supervisorEmail);
      if (!signature.email.trim()) {
        signature = {
          ...signature,
          email: supervisorEmail,
          name: signature.name.trim() || resource.supervisor.trim(),
        };
      }
      if (
        !signature.name.trim() &&
        !signature.position.trim() &&
        !signature.phone.trim()
      ) {
        signature = {
          ...createEmptyEmailSignature(),
          name: resource.supervisor.trim() || supervisorEmail,
          email: supervisorEmail,
        };
      }

      try {
        await sendGraphAppMail({
          from: supervisorEmail,
          to,
          cc,
          subject: content.subject,
          bodyText: content.body,
          portalHref: content.portalHref,
          signature,
          publicOrigin,
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
