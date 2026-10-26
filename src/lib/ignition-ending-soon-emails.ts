import {
  collectNonBusinessDates,
  ensureWorkingDay,
} from "@/lib/business-days";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import type { EmailTemplate } from "@/lib/email-templates";
import type { HolidaysStore } from "@/lib/holidays";
import type { EndingSoonEmailStatusLabel } from "@/lib/ignition-ending-soon-email-ui";
import { ENDING_SOON_EMAIL_SENDERS } from "@/lib/ignition-ending-soon-email-ui";
import { isValidEmail, normalizeEmail } from "@/lib/portal-users";
import { USER_TIME_ZONE, zonedWallTimeToUtc } from "@/lib/time-zones";

/** Team Ending Soon deferred send hour in Asia/Karachi (21 = 9:00 PM). */
export const ENDING_SOON_TEAM_EMAIL_SEND_HOUR = 21;

/** Local copy — avoid importing ignition/proposals (pulls fs into the client). */
function normalizeMatchName(value: string): string {
  return value
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\b(llc|inc|ltd|corp|co|plc|pllc|pc)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * HARD STOP: Ending Soon is status/logic only. Never call Graph / sendGraphMail.
 */
export const ENDING_SOON_EMAILS_DRY_RUN = true;

/**
 * For proposals whose endsOn year is after the current year (e.g. 2027),
 * plan the team email for the next working day at 9:00 PM Pakistan time
 * (instead of waiting until 2 months before expiry).
 */
export const ENDING_SOON_TEAM_EMAIL_SCHEDULE_NOW_FOR_NEXT_YEAR = true;

export type EndingSoonEmailStatus = "scheduled" | "sent";
export type EndingSoonEmailAudience = "team" | "client";

export type EndingSoonEmailEntry = {
  /** Proposal slug. */
  proposalSlug: string;
  audience: EndingSoonEmailAudience;
  clientName: string;
  referenceNumber: string;
  endsOn: string;
  status: EndingSoonEmailStatus;
  scheduledFor: string;
  sentAt: string;
  recipients: string[];
};

export type EndingSoonEmailsStore = {
  dryRun: boolean;
  entries: EndingSoonEmailEntry[];
};

export type EndingSoonProposalInput = {
  slug: string;
  clientSlug: string;
  clientName: string;
  name: string;
  referenceNumber: string;
  endsOn: string;
};

export type EndingSoonEmailPlan = {
  proposalSlug: string;
  audience: EndingSoonEmailAudience;
  clientName: string;
  referenceNumber: string;
  endsOn: string;
  subject: string;
  body: string;
  /** Deferred Graph send time (ISO). Null → send immediately. */
  sendAt: string | null;
  /** BCC list (team = portal users; client = that client's emails only). */
  bcc: string[];
  allowedFrom: readonly string[];
  markAs: EndingSoonEmailStatus;
};

export type { EndingSoonEmailStatusLabel } from "@/lib/ignition-ending-soon-email-ui";
export {
  ENDING_SOON_EMAIL_SENDER,
  ENDING_SOON_EMAIL_SENDERS,
  endingSoonEmailSendersLabel,
  isEndingSoonEmailSender,
} from "@/lib/ignition-ending-soon-email-ui";

export const ENDING_SOON_TEAM_EMAIL_TEMPLATE_ID = "template-19";
export const ENDING_SOON_CLIENT_EMAIL_TEMPLATE_ID = "template-20";

/**
 * Team notice: two full calendar months before endsOn (weekends/holidays count
 * in that span). If the resulting day is a weekend or holiday, schedule on the
 * next working day after it.
 */
export const ENDING_SOON_TEAM_EMAIL_LEAD_MONTHS = 2;

/**
 * Client notice: one full calendar month before endsOn (weekends/holidays count).
 * Weekend/holiday → next working day. Same 9:00 PM Pakistan time.
 */
export const ENDING_SOON_CLIENT_EMAIL_LEAD_MONTHS = 1;

export function createEmptyEndingSoonEmailsStore(): EndingSoonEmailsStore {
  return { dryRun: true, entries: [] };
}

export function getEndingSoonEmailEntry(
  store: EndingSoonEmailsStore,
  proposalSlug: string,
  audience: EndingSoonEmailAudience,
): EndingSoonEmailEntry | null {
  return (
    store.entries.find(
      (item) =>
        item.proposalSlug === proposalSlug && item.audience === audience,
    ) ?? null
  );
}

function profileEmails(profile: ClientProfileEntry): string[] {
  const emails: string[] = [];
  for (const raw of [
    profile.pocEmail,
    profile.documentRequestEmail,
    profile.financialsEmail,
    profile.ownerEmail,
  ]) {
    const email = normalizeEmail(String(raw ?? ""));
    if (email && isValidEmail(email)) {
      emails.push(email);
    }
  }
  return [...new Set(emails)];
}

export function matchEndingSoonProposalToProfile(
  proposal: EndingSoonProposalInput,
  profiles: ClientProfileEntry[],
): ClientProfileEntry | null {
  const slug = proposal.clientSlug.trim();
  if (slug) {
    const bySlug = profiles.find(
      (profile) => profile.ignitionClientSlug.trim() === slug,
    );
    if (bySlug) {
      return bySlug;
    }
  }

  const target = normalizeMatchName(proposal.clientName);
  if (!target) {
    return null;
  }
  for (const profile of profiles) {
    const names = [profile.legalName, profile.operatingName]
      .map((value) => normalizeMatchName(value))
      .filter(Boolean);
    if (names.includes(target)) {
      return profile;
    }
  }
  return null;
}

export function collectEndingSoonClientEmails(
  proposal: EndingSoonProposalInput,
  profiles: ClientProfileEntry[],
  ignitionClientEmail?: string,
): string[] {
  const profile = matchEndingSoonProposalToProfile(proposal, profiles);
  if (profile) {
    return profileEmails(profile);
  }
  const fallback = normalizeEmail(String(ignitionClientEmail ?? ""));
  return fallback && isValidEmail(fallback) ? [fallback] : [];
}

export function collectEndingSoonTeamEmails(
  portalUsers: Array<{ email?: string }>,
): string[] {
  const emails = new Set<string>();
  for (const user of portalUsers) {
    const email = normalizeEmail(String(user.email ?? ""));
    if (email && isValidEmail(email)) {
      emails.add(email);
    }
  }
  return [...emails].sort((a, b) => a.localeCompare(b));
}

function formatExpiryDisplay(endsOn: string): string {
  if (!/^\d{4}-\d{2}-\d{2}/.test(endsOn.trim())) {
    return endsOn.trim() || "—";
  }
  const date = new Date(`${endsOn.trim().slice(0, 10)}T00:00:00`);
  if (Number.isNaN(date.getTime())) {
    return endsOn.trim();
  }
  return date.toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

function applyEndingSoonPlaceholders(
  text: string,
  proposal: EndingSoonProposalInput,
): string {
  const expiry = formatExpiryDisplay(proposal.endsOn);
  return text
    .replaceAll("[Client Name]", proposal.clientName || "Client")
    .replaceAll(
      "[Proposal Reference]",
      proposal.referenceNumber || proposal.slug || "—",
    )
    .replaceAll("[Proposal Name]", proposal.name || "Proposal")
    .replaceAll("[Expiry Date]", expiry);
}

function parseEndsOnLocal(endsOn: string): Date | null {
  if (!/^\d{4}-\d{2}-\d{2}/.test(endsOn.trim())) {
    return null;
  }
  const [y, m, d] = endsOn.trim().slice(0, 10).split("-").map(Number);
  if (!y || !m || !d) {
    return null;
  }
  const date = new Date(y, m - 1, d);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return date;
}

function wallTimeIso(date: Date): string {
  return zonedWallTimeToUtc(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    ENDING_SOON_TEAM_EMAIL_SEND_HOUR,
    0,
    USER_TIME_ZONE,
  ).toISOString();
}

function subtractCalendarMonths(end: Date, months: number): Date {
  const target = new Date(
    end.getFullYear(),
    end.getMonth() - months,
    end.getDate(),
  );
  // Month-end clamp (e.g. Mar 31 → Jan/Feb last day).
  if (target.getDate() !== end.getDate()) {
    target.setDate(0);
  }
  return target;
}

/** Next working day at 9:00 PM Asia/Karachi that is still in the future. */
function ensureFutureNinePmPkt(
  startDay: Date,
  excludedDates: Set<string>,
  now: Date,
): string | null {
  let day = ensureWorkingDay(
    new Date(startDay.getFullYear(), startDay.getMonth(), startDay.getDate()),
    excludedDates,
  );
  const soonest = now.getTime() + 60_000;
  for (let step = 0; step < 30; step += 1) {
    const iso = wallTimeIso(day);
    if (Date.parse(iso) >= soonest) {
      return iso;
    }
    const next = new Date(day.getFullYear(), day.getMonth(), day.getDate() + 1);
    day = ensureWorkingDay(next, excludedDates);
  }
  return null;
}

/**
 * Team: two calendar months before endsOn (weekends/holidays included in the
 * span). If that day is a weekend or holiday, move to the next working day.
 * Always at 9:00 PM Asia/Karachi (deferred in Outlook when live).
 * When {@link ENDING_SOON_TEAM_EMAIL_SCHEDULE_NOW_FOR_NEXT_YEAR} is on and the
 * expiry year is after this year, schedule for the next working day at 9:00 PM PKT.
 */
export function plannedTeamSendAtForEndsOn(
  endsOn: string,
  excludedDates: Set<string> = new Set(),
  now: Date = new Date(),
): string | null {
  const end = parseEndsOnLocal(endsOn);
  if (!end) {
    return null;
  }
  let target: Date;
  if (
    ENDING_SOON_TEAM_EMAIL_SCHEDULE_NOW_FOR_NEXT_YEAR &&
    end.getFullYear() > now.getFullYear()
  ) {
    target = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  } else {
    target = subtractCalendarMonths(end, ENDING_SOON_TEAM_EMAIL_LEAD_MONTHS);
  }
  return ensureFutureNinePmPkt(target, excludedDates, now);
}

/**
 * Client: one calendar month before endsOn at 9:00 PM Pakistan time
 * (weekends/holidays count; weekend/holiday → next working day).
 */
export function plannedClientSendAtForEndsOn(
  endsOn: string,
  excludedDates: Set<string> = new Set(),
  now: Date = new Date(),
): string | null {
  const end = parseEndsOnLocal(endsOn);
  if (!end) {
    return null;
  }
  const target = subtractCalendarMonths(
    end,
    ENDING_SOON_CLIENT_EMAIL_LEAD_MONTHS,
  );
  return ensureFutureNinePmPkt(target, excludedDates, now);
}

/** Holiday dates for a matched client profile (Pakistan + client country). */
export function collectEndingSoonClientExcludedDates(
  holidays: HolidaysStore,
  profile: ClientProfileEntry | null,
  endsOn: string,
): Set<string> {
  const end = parseEndsOnLocal(endsOn);
  if (!end) {
    return new Set();
  }
  const years = new Set<number>([
    end.getFullYear(),
    end.getFullYear() - 1,
    end.getFullYear() + 1,
  ]);
  const excluded = new Set<string>();
  for (const year of years) {
    for (const date of collectNonBusinessDates(
      holidays,
      profile?.country,
      profile?.state ?? "",
      year,
    )) {
      excluded.add(date);
    }
  }
  return excluded;
}

export function fillEndingSoonEmailTemplate(
  template: EmailTemplate | null | undefined,
  proposal: EndingSoonProposalInput,
): { subject: string; body: string } {
  const subject = applyEndingSoonPlaceholders(
    template?.subject ||
      "Reminder: [Client Name] proposal expires on [Expiry Date]",
    proposal,
  );
  const body = applyEndingSoonPlaceholders(
    template?.body ||
      "This proposal is going to expire on [Expiry Date].\n\nThank you.",
    proposal,
  );
  return { subject, body };
}

function normalizeEntry(value: unknown): EndingSoonEmailEntry | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<EndingSoonEmailEntry>;
  const proposalSlug = String(record.proposalSlug ?? "").trim();
  if (!proposalSlug) {
    return null;
  }
  const audience: EndingSoonEmailAudience =
    record.audience === "client" ? "client" : "team";
  const recipients = Array.isArray(record.recipients)
    ? record.recipients
        .map((item) => normalizeEmail(String(item ?? "")))
        .filter((email) => email && isValidEmail(email))
    : [];
  const status: EndingSoonEmailStatus =
    record.status === "sent" ? "sent" : "scheduled";
  return {
    proposalSlug,
    audience,
    clientName: String(record.clientName ?? "").trim(),
    referenceNumber: String(record.referenceNumber ?? "").trim(),
    endsOn: String(record.endsOn ?? "").trim(),
    status,
    scheduledFor: String(record.scheduledFor ?? "").trim(),
    sentAt: String(record.sentAt ?? "").trim(),
    recipients,
  };
}

export function normalizeEndingSoonEmailsStore(
  value: unknown,
): EndingSoonEmailsStore {
  const record =
    typeof value === "object" && value !== null
      ? (value as { dryRun?: unknown; entries?: unknown })
      : {};
  const wasDryRun = record.dryRun === true;
  const entries = Array.isArray(record.entries)
    ? record.entries
        .map(normalizeEntry)
        .filter((item): item is EndingSoonEmailEntry => Boolean(item))
    : [];
  // Drop prior dry-run-only rows when switching into live send mode.
  const clearForLive = wasDryRun && !ENDING_SOON_EMAILS_DRY_RUN;
  return {
    dryRun: ENDING_SOON_EMAILS_DRY_RUN,
    entries: clearForLive ? [] : entries,
  };
}

export function upsertEndingSoonEmailEntry(
  store: EndingSoonEmailsStore,
  entry: EndingSoonEmailEntry,
): EndingSoonEmailsStore {
  const without = store.entries.filter(
    (item) =>
      !(
        item.proposalSlug === entry.proposalSlug &&
        item.audience === entry.audience
      ),
  );
  return {
    dryRun: ENDING_SOON_EMAILS_DRY_RUN,
    entries: [...without, entry],
  };
}

export function buildEndingSoonEmailStatusBySlug(
  proposals: EndingSoonProposalInput[],
  store: EndingSoonEmailsStore,
  audience: EndingSoonEmailAudience,
): Record<string, EndingSoonEmailStatusLabel> {
  const bySlug = new Map<string, EndingSoonEmailEntry>();
  for (const entry of store.entries) {
    if (entry.audience === audience) {
      bySlug.set(entry.proposalSlug, entry);
    }
  }
  const result: Record<string, EndingSoonEmailStatusLabel> = {};
  for (const proposal of proposals) {
    const entry = bySlug.get(proposal.slug);
    if (!entry) {
      result[proposal.slug] = "—";
    } else if (entry.status === "sent" && !ENDING_SOON_EMAILS_DRY_RUN) {
      result[proposal.slug] = "Sent";
    } else {
      result[proposal.slug] = "Scheduled";
    }
  }
  return result;
}

export type BuildEndingSoonEmailPlansInput = {
  proposals: EndingSoonProposalInput[];
  teamEmails: string[];
  profiles: ClientProfileEntry[];
  store: EndingSoonEmailsStore;
  /** Ignition client slug → email fallback. */
  ignitionEmailsByClientSlug?: Record<string, string>;
  teamTemplate?: EmailTemplate | null;
  clientTemplate?: EmailTemplate | null;
  holidays?: HolidaysStore | null;
  now?: Date;
};

function resolveMarkAs(
  sendAtIso: string | null,
  now: Date,
): { markAs: EndingSoonEmailStatus; sendAt: string | null } {
  // Status/logic only — always Scheduled, never Sent / never Graph.
  return {
    markAs: "scheduled",
    sendAt: sendAtIso || now.toISOString(),
  };
}

/** Persist Scheduled entries from plans (used in dry-run; no Graph send). */
export function applyEndingSoonEmailPlansToStore(
  store: EndingSoonEmailsStore,
  plans: EndingSoonEmailPlan[],
): EndingSoonEmailsStore {
  let next = { ...store, dryRun: ENDING_SOON_EMAILS_DRY_RUN };
  for (const plan of plans) {
    next = upsertEndingSoonEmailEntry(
      next,
      entryFromEndingSoonPlan(plan, plan.markAs),
    );
  }
  return next;
}

/**
 * Build team + client schedule plans (status/logic only — never Graph).
 * In dry-run, rebuilds Scheduled for every matching proposal.
 */
export function buildEndingSoonEmailPlans(
  input: BuildEndingSoonEmailPlansInput,
): EndingSoonEmailPlan[] {
  const plans: EndingSoonEmailPlan[] = [];
  const holidays = input.holidays ?? null;
  const now = input.now ?? new Date();

  for (const proposal of input.proposals) {
    if (!proposal.slug.trim() || !proposal.endsOn.trim()) {
      continue;
    }
    const matchedProfile = matchEndingSoonProposalToProfile(
      proposal,
      input.profiles,
    );
    const excluded = holidays
      ? collectEndingSoonClientExcludedDates(
          holidays,
          matchedProfile,
          proposal.endsOn,
        )
      : new Set<string>();

    if (input.teamEmails.length > 0) {
      const existingTeam = getEndingSoonEmailEntry(
        input.store,
        proposal.slug,
        "team",
      );
      if (
        !ENDING_SOON_EMAILS_DRY_RUN &&
        (existingTeam?.status === "sent" ||
          existingTeam?.status === "scheduled")
      ) {
        // live path reserved; dry-run always rebuilds
      } else {
        const filled = fillEndingSoonEmailTemplate(
          input.teamTemplate,
          proposal,
        );
        const rawSendAt = plannedTeamSendAtForEndsOn(
          proposal.endsOn,
          excluded,
          now,
        );
        const { markAs, sendAt } = resolveMarkAs(rawSendAt, now);
        plans.push({
          proposalSlug: proposal.slug,
          audience: "team",
          clientName: proposal.clientName,
          referenceNumber: proposal.referenceNumber,
          endsOn: proposal.endsOn,
          subject: filled.subject,
          body: filled.body,
          sendAt,
          bcc: input.teamEmails,
          allowedFrom: ENDING_SOON_EMAIL_SENDERS,
          markAs,
        });
      }
    }

    const clientEmails = collectEndingSoonClientEmails(
      proposal,
      input.profiles,
      input.ignitionEmailsByClientSlug?.[proposal.clientSlug],
    );
    if (clientEmails.length > 0) {
      const existingClient = getEndingSoonEmailEntry(
        input.store,
        proposal.slug,
        "client",
      );
      if (
        !ENDING_SOON_EMAILS_DRY_RUN &&
        (existingClient?.status === "sent" ||
          existingClient?.status === "scheduled")
      ) {
        // live path reserved
      } else {
        const filled = fillEndingSoonEmailTemplate(
          input.clientTemplate,
          proposal,
        );
        const rawSendAt = plannedClientSendAtForEndsOn(
          proposal.endsOn,
          excluded,
          now,
        );
        const { markAs, sendAt } = resolveMarkAs(rawSendAt, now);
        plans.push({
          proposalSlug: proposal.slug,
          audience: "client",
          clientName: proposal.clientName,
          referenceNumber: proposal.referenceNumber,
          endsOn: proposal.endsOn,
          subject: filled.subject,
          body: filled.body,
          sendAt,
          bcc: clientEmails,
          allowedFrom: ENDING_SOON_EMAIL_SENDERS,
          markAs,
        });
      }
    }
  }

  return plans;
}

export function entryFromEndingSoonPlan(
  plan: EndingSoonEmailPlan,
  status: EndingSoonEmailStatus,
): EndingSoonEmailEntry {
  const nowIso = new Date().toISOString();
  return {
    proposalSlug: plan.proposalSlug,
    audience: plan.audience,
    clientName: plan.clientName,
    referenceNumber: plan.referenceNumber,
    endsOn: plan.endsOn,
    status: ENDING_SOON_EMAILS_DRY_RUN ? "scheduled" : status,
    scheduledFor: plan.sendAt || nowIso,
    sentAt: "",
    recipients: plan.bcc,
  };
}
