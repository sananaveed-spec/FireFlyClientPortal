import { isGraphAppOnlyConfigured } from "@/lib/azure-graph-app";
import { findProfileForResourceClient } from "@/lib/client-management";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import {
  readClientProfiles,
  readEmailSignatureForUser,
  readEmailTemplates,
  readPortalUsersFile,
  readResourceClients,
  readTaxTaskAutoEmailStatusFile,
  writeTaxTaskAutoEmailStatusFile,
} from "@/lib/data-store";
import {
  applyPeriodPlaceholder,
  CORPORATE_INCOME_TAX_TEMPLATE_ID,
  GST_HST_RETURN_TEMPLATE_ID,
  PERSONAL_INCOME_TAX_TEMPLATE_ID,
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
import { parsePortalUsers, type PortalUser } from "@/lib/portal-users";
import { normalizeAnnualTasks } from "@/lib/annual-tasks";
import { normalizeQuarterlyTasks } from "@/lib/quarterly-tasks";
import { sendGraphAppMail } from "@/lib/send-graph-app-mail";
import { resolveStaffEmail } from "@/lib/staff-assign";
import {
  getTaxTaskAutoEmailStatus,
  markTaxTaskAutoEmailError,
  markTaxTaskAutoEmailSent,
  normalizeTaxTaskAutoEmailStatusStore,
  type TaxTaskAutoEmailKind,
} from "@/lib/tax-task-auto-email-status";
import {
  getCalendarYmdInTimeZone,
  getClientNineAmSchedule,
  resolveClientTimeZone,
} from "@/lib/time-zones";

const SCHEDULE_EMAIL_OWNER_CC_FALLBACK = "ayesha@fireflyaccountancy.com";
const EPS_SCHEDULE_EMAIL_CC = "thomas@epsfresno.com";

/** Same window as BD auto-send (covers multi-TZ cron passes). */
const SEND_WINDOW_MS = 14 * 60 * 60 * 1000;

const TAX_TASK_SPECS: Array<{
  kind: TaxTaskAutoEmailKind;
  taskKind: string;
  list: "annual" | "quarterly";
  templateId: string;
  scheduleId: string;
}> = [
  {
    kind: "corporate-income-tax",
    taskKind: "income_tax_compliance:corporate-return",
    list: "annual",
    templateId: CORPORATE_INCOME_TAX_TEMPLATE_ID,
    scheduleId: "tax-corporate-income-tax",
  },
  {
    kind: "personal-income-tax",
    taskKind: "income_tax_compliance:personal-return",
    list: "annual",
    templateId: PERSONAL_INCOME_TAX_TEMPLATE_ID,
    scheduleId: "tax-personal-income-tax",
  },
  {
    kind: "gst-hst-return",
    taskKind: "gst_hst_tax_compliance:return-filing",
    list: "quarterly",
    templateId: GST_HST_RETURN_TEMPLATE_ID,
    scheduleId: "tax-gst-hst-return",
  },
];

export type TaxTaskAutoEmailRunResult = {
  ok: boolean;
  checkedAt: string;
  sent: Array<{
    clientId: string;
    clientName: string;
    scheduleId: string;
    from: string;
    to: string;
    lastSentAt: string;
  }>;
  skipped: Array<{
    clientId: string;
    clientName: string;
    scheduleId: string;
    reason: string;
  }>;
  errors: Array<{
    clientId: string;
    clientName: string;
    scheduleId: string;
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

function formatDueDateLabel(ymd: string): string {
  const match = ymd.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return ymd.trim() || "—";
  }
  const date = new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
  );
  if (Number.isNaN(date.getTime())) {
    return ymd;
  }
  return date.toLocaleDateString(undefined, {
    weekday: "long",
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function publicPortalOrigin(): string {
  return (
    process.env.IGNITION_PUBLIC_BASE_URL?.trim() ||
    process.env.NEXT_PUBLIC_APP_URL?.trim() ||
    ""
  );
}

function buildTaxEmailContent(input: {
  profile: ClientProfileEntry;
  template: EmailTemplate;
  dueDateYmd: string;
  now: Date;
}): { subject: string; body: string; portalHref: string | null } {
  const { profile, template, dueDateYmd, now } = input;
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
  const dueDateLabel = formatDueDateLabel(dueDateYmd);

  function fill(text: string) {
    return applyPeriodPlaceholder(
      applyPortalLinkPlaceholders(
        text
          .replace(/\[Client Name\]/gi, clientName)
          .replace(/\[Client\]/gi, clientName)
          .replace(/\[Contact First Name\]/gi, contactName)
          .replace(/\[Due Date\]/gi, dueDateLabel)
          .replace(/\[Deadline\]/gi, dueDateLabel),
        portalLink,
        now,
      ),
      now,
    );
  }

  const subject = fill(template.subject?.trim() || "");
  const body = fill(
    stripLegacySignatureClosing(template.body ?? "").trimEnd(),
  ).replace(/\n{3,}/g, "\n\n");

  return {
    subject,
    body,
    portalHref: resolvePortalHref(portalLink, now) || null,
  };
}

/**
 * Auto-send Corporate / Personal income tax and GST/HST return emails when
 * the matching Annual/Quarterly task due date is today (9:00 AM client time).
 * To / CC / From match BD schedule emails (POC + owner/group/EPS CC).
 */
export async function runTaxTaskAutoEmails(options?: {
  force?: boolean;
  now?: Date;
}): Promise<TaxTaskAutoEmailRunResult> {
  const now = options?.now ?? new Date();
  const force = Boolean(options?.force);
  const result: TaxTaskAutoEmailRunResult = {
    ok: true,
    checkedAt: now.toISOString(),
    sent: [],
    skipped: [],
    errors: [],
  };

  if (!isGraphAppOnlyConfigured()) {
    result.ok = false;
    result.errors.push({
      clientId: "",
      clientName: "",
      scheduleId: "tax-corporate-income-tax",
      error:
        "AZURE_CLIENT_SECRET is not set (required for app-only Mail.Send).",
    });
    return result;
  }

  const [profiles, resources, templates, usersRaw, statusRaw] =
    await Promise.all([
      readClientProfiles(),
      readResourceClients(),
      readEmailTemplates(),
      readPortalUsersFile(),
      readTaxTaskAutoEmailStatusFile(),
    ]);

  const users = parsePortalUsers(usersRaw);
  let statusStore = normalizeTaxTaskAutoEmailStatusStore(statusRaw);
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

    const annualTasks = normalizeAnnualTasks(profile.annualTasks);
    const quarterlyTasks = normalizeQuarterlyTasks(profile.quarterlyTasks);

    for (const spec of TAX_TASK_SPECS) {
      const tasks = spec.list === "annual" ? annualTasks : quarterlyTasks;
      const task = tasks.find((item) => item.kind === spec.taskKind);
      if (!task) {
        continue;
      }
      const dueDate = task.dueDate.trim();
      if (!dueDate) {
        continue;
      }
      if (dueDate !== todayYmd) {
        continue;
      }

      const existing = getTaxTaskAutoEmailStatus(
        statusStore,
        profile.id,
        spec.kind,
        dueDate,
      );
      if (!force && existing?.emailSent) {
        continue;
      }

      const supervisorEmail = resolveStaffEmail(resource.supervisor, users);
      if (!supervisorEmail) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
          reason: "No supervisor email on Resource Allocation.",
        });
        continue;
      }

      const to = profile.pocEmail.trim();
      if (!to) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
          reason: "No POC Email on Client Profile.",
        });
        continue;
      }

      // Interpret due date as a local calendar day for 9:00 AM scheduling.
      const [y, m, d] = dueDate.split("-").map(Number);
      const dueLocal = new Date(y!, (m ?? 1) - 1, d ?? 1);
      const sendSchedule = getClientNineAmSchedule(dueLocal, profile.timeZone);
      if (!sendSchedule) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
          reason: "Client time zone missing or invalid.",
        });
        continue;
      }

      const sendAt = sendSchedule.instant;
      if (now.getTime() < sendAt.getTime()) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
          reason: `Send time not reached yet (${sendAt.toISOString()}).`,
        });
        continue;
      }
      if (now.getTime() > sendAt.getTime() + SEND_WINDOW_MS) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
          reason: `Outside send window (sendAt ${sendAt.toISOString()}).`,
        });
        continue;
      }

      const template = findTemplateById(templates, spec.templateId);
      if (!template) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
          reason: `Email template ${spec.templateId} not found.`,
        });
        continue;
      }

      const content = buildTaxEmailContent({
        profile,
        template,
        dueDateYmd: dueDate,
        now,
      });
      if (!content.subject.trim()) {
        result.skipped.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
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
      const supervisorLabel = resource.supervisor.trim();
      const supervisorDisplayName = supervisorLabel.includes("@")
        ? ""
        : supervisorLabel;
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
          bcc: "",
          subject: content.subject,
          bodyText: content.body,
          portalHref: content.portalHref,
          signature,
          publicOrigin,
        });
        const lastSentAt = now.toISOString();
        statusStore = markTaxTaskAutoEmailSent(statusStore, {
          profileId: profile.id,
          kind: spec.kind,
          dueDate,
          lastSentAt,
          fromEmail: supervisorEmail,
        });
        result.sent.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
          from: supervisorEmail,
          to,
          lastSentAt,
        });
      } catch (sendError) {
        const message =
          sendError instanceof Error
            ? sendError.message
            : "Failed to send tax task auto email.";
        statusStore = markTaxTaskAutoEmailError(statusStore, {
          profileId: profile.id,
          kind: spec.kind,
          dueDate,
          lastError: message,
          fromEmail: supervisorEmail,
        });
        result.errors.push({
          clientId: resource.id,
          clientName,
          scheduleId: spec.scheduleId,
          error: message,
        });
        result.ok = false;
      }
    }
  }

  await writeTaxTaskAutoEmailStatusFile(statusStore);
  return result;
}
