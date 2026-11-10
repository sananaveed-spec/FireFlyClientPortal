import { isGraphAppOnlyConfigured } from "@/lib/azure-graph-app";
import {
  readClientProfiles,
  readEmailSignatureForUser,
  readEmailTemplates,
  readHolidays,
  readPakistanHolidayEmailsFile,
  writePakistanHolidayEmailsFile,
} from "@/lib/data-store";
import { normalizeEmailTemplates } from "@/lib/email-templates";
import { normalizeHolidays } from "@/lib/holidays";
import {
  buildPakistanHolidayEmailPlans,
  collectPakistanHolidayClientRecipientGroups,
  normalizePakistanHolidayEmailsStore,
  PAKISTAN_HOLIDAY_CLIENT_EMAIL_TEMPLATE_ID,
  PAKISTAN_HOLIDAY_EMAIL_SENDER,
  PAKISTAN_HOLIDAY_EMAIL_TEMPLATE_ID,
  upsertPakistanHolidayEmailEntry,
  type PakistanHolidayEmailEntry,
  type PakistanHolidayEmailPlan,
  type PakistanHolidayEmailsStore,
} from "@/lib/pakistan-holiday-emails";
import { readPortalUsers } from "@/lib/portal-users-store";
import { sendGraphAppMail } from "@/lib/send-graph-app-mail";

/**
 * How long after the planned 9:00 AM (Pakistan / client TZ) a cron may still
 * send. Keep tight so mail goes out near 9:00 AM, not hours later.
 * Pair with an hourly holiday-only cron.
 */
const SEND_WINDOW_MS = 90 * 60 * 1000;

export type PakistanHolidayAutoEmailRunResult = {
  ok: boolean;
  checkedAt: string;
  queued: number;
  sent: Array<{
    audience: string;
    holidayName: string;
    holidayDate: string;
    recipients: number;
  }>;
  skipped: Array<{ reason: string }>;
  errors: Array<{ holidayName: string; audience: string; error: string }>;
};

function portalPublicOrigin(): string {
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

function pakistanHolidaysFromStore() {
  return readHolidays().then((raw) => {
    const holidays = normalizeHolidays(raw);
    const pakistan = holidays.countries.find(
      (country) => country.id === "pakistan",
    );
    return pakistan?.holidays ?? [];
  });
}

function planIsDue(
  plan: PakistanHolidayEmailPlan,
  now: Date,
  force: boolean,
): boolean {
  if (!plan.sendAt) {
    return true;
  }
  const sendAtMs = Date.parse(plan.sendAt);
  if (!Number.isFinite(sendAtMs)) {
    return false;
  }
  if (sendAtMs > now.getTime()) {
    return false;
  }
  if (force) {
    return true;
  }
  return now.getTime() - sendAtMs <= SEND_WINDOW_MS;
}

function planIsFuture(plan: PakistanHolidayEmailPlan, now: Date): boolean {
  if (!plan.sendAt) {
    return false;
  }
  const sendAtMs = Date.parse(plan.sendAt);
  return Number.isFinite(sendAtMs) && sendAtMs > now.getTime();
}

function entryFromPlan(
  plan: PakistanHolidayEmailPlan,
  status: "scheduled" | "sent",
  now: Date,
): PakistanHolidayEmailEntry {
  return {
    groupKey: plan.groupKey,
    audience: plan.audience,
    holidayDate: plan.holidayDate,
    holidayDates: plan.holidayDates,
    holidayIds: plan.holidayIds,
    holidayName: plan.holidayName,
    status,
    scheduledFor: plan.sendAt || now.toISOString(),
    sentAt: status === "sent" ? now.toISOString() : "",
  };
}

/**
 * Unattended Pakistan holiday emails (team day-before + client 7 business days).
 * Queues upcoming rows as Scheduled, and sends due ones via app-only Mail.Send
 * as ayesha@… — same pattern as BD auto-cron (no one needs to open the portal).
 */
export async function runPakistanHolidayAutoEmails(options?: {
  force?: boolean;
  now?: Date;
}): Promise<PakistanHolidayAutoEmailRunResult> {
  const now = options?.now ?? new Date();
  const force = Boolean(options?.force);
  const checkedAt = now.toISOString();
  const result: PakistanHolidayAutoEmailRunResult = {
    ok: true,
    checkedAt,
    queued: 0,
    sent: [],
    skipped: [],
    errors: [],
  };

  if (!isGraphAppOnlyConfigured()) {
    result.ok = false;
    result.skipped.push({
      reason:
        "AZURE_CLIENT_SECRET is not set (required for app-only Mail.Send).",
    });
    return result;
  }

  const [pakistanHolidays, templatesRaw, emailsRaw, users, profiles] =
    await Promise.all([
      pakistanHolidaysFromStore(),
      readEmailTemplates(),
      readPakistanHolidayEmailsFile(),
      readPortalUsers(),
      readClientProfiles(),
    ]);

  if (pakistanHolidays.length === 0) {
    result.skipped.push({ reason: "No Pakistan holidays configured." });
    return result;
  }

  const templates = normalizeEmailTemplates(templatesRaw);
  const teamTemplate = templates.templates.find(
    (item) => item.id === PAKISTAN_HOLIDAY_EMAIL_TEMPLATE_ID,
  );
  const clientTemplate = templates.templates.find(
    (item) => item.id === PAKISTAN_HOLIDAY_CLIENT_EMAIL_TEMPLATE_ID,
  );
  if (!teamTemplate || !clientTemplate) {
    result.ok = false;
    result.skipped.push({
      reason: "Pakistan holiday email templates (16/17) are missing.",
    });
    return result;
  }

  let store = normalizePakistanHolidayEmailsStore(emailsRaw);

  /** Build plans while treating already-scheduled rows as still eligible when due. */
  const storeForPlanning: PakistanHolidayEmailsStore = {
    entries: store.entries.filter((entry) => entry.status === "sent"),
  };

  const teamRecipients = users
    .map((user) => user.email.trim())
    .filter(Boolean);
  const clientRecipientGroups =
    collectPakistanHolidayClientRecipientGroups(profiles);

  const teamPlans = buildPakistanHolidayEmailPlans(
    pakistanHolidays,
    storeForPlanning,
    { subject: teamTemplate.subject, body: teamTemplate.body },
    { audience: "team", now },
  );
  const clientPlans = buildPakistanHolidayEmailPlans(
    pakistanHolidays,
    storeForPlanning,
    { subject: clientTemplate.subject, body: clientTemplate.body },
    {
      audience: "client",
      recipientGroups: clientRecipientGroups,
      now,
    },
  );

  const allPlans = [...teamPlans, ...clientPlans];
  if (allPlans.length === 0) {
    result.skipped.push({ reason: "No upcoming Pakistan holiday emails." });
    return result;
  }

  const from = PAKISTAN_HOLIDAY_EMAIL_SENDER;
  const signature = await readEmailSignatureForUser(from);
  const publicOrigin = portalPublicOrigin();

  for (const plan of allPlans) {
    const existing = store.entries.find(
      (entry) =>
        entry.groupKey === plan.groupKey && entry.audience === plan.audience,
    );
    if (existing?.status === "sent") {
      continue;
    }

    if (planIsFuture(plan, now)) {
      if (!existing || existing.status !== "scheduled") {
        store = upsertPakistanHolidayEmailEntry(
          store,
          entryFromPlan(plan, "scheduled", now),
        );
        result.queued += 1;
      }
      continue;
    }

    if (!planIsDue(plan, now, force)) {
      result.skipped.push({
        reason: `${plan.audience} “${plan.holidayName}” missed the send window.`,
      });
      continue;
    }

    const recipients =
      plan.audience === "team"
        ? teamRecipients
        : (plan.recipients ?? []).filter(Boolean);
    if (recipients.length === 0) {
      result.skipped.push({
        reason: `${plan.audience} “${plan.holidayName}” has no recipients.`,
      });
      continue;
    }

    try {
      await sendGraphAppMail({
        from,
        to: from,
        bcc: recipients.join(", "),
        subject: plan.subject,
        bodyText: plan.body,
        highlights: [plan.holidayName],
        signature,
        publicOrigin,
      });
      store = upsertPakistanHolidayEmailEntry(
        store,
        entryFromPlan(plan, "sent", now),
      );
      result.sent.push({
        audience: plan.audience,
        holidayName: plan.holidayName,
        holidayDate: plan.holidayDate,
        recipients: recipients.length,
      });
    } catch (error) {
      result.ok = false;
      result.errors.push({
        audience: plan.audience,
        holidayName: plan.holidayName,
        error:
          error instanceof Error
            ? error.message
            : "Failed to send Pakistan holiday email.",
      });
    }
  }

  await writePakistanHolidayEmailsFile(store);
  return result;
}
