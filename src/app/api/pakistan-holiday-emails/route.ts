import { NextResponse } from "next/server";
import {
  readClientProfiles,
  readEmailTemplates,
  readHolidays,
  readPakistanHolidayEmailsFile,
  writePakistanHolidayEmailsFile,
} from "@/lib/data-store";
import { normalizeEmailTemplates } from "@/lib/email-templates";
import { normalizeHolidays } from "@/lib/holidays";
import {
  buildPakistanHolidayEmailPlans,
  buildPakistanHolidayEmailPreview,
  buildPakistanHolidayEmailStatusByHolidayId,
  collectPakistanHolidayClientRecipientEmails,
  collectPakistanHolidayClientRecipientGroups,
  isPakistanHolidayEmailSender,
  normalizePakistanHolidayEmailsStore,
  PAKISTAN_HOLIDAY_CLIENT_EMAIL_TEMPLATE_ID,
  PAKISTAN_HOLIDAY_EMAIL_TEMPLATE_ID,
  pakistanHolidayEmailSendersLabel,
  promoteDueScheduledEntries,
  upsertPakistanHolidayEmailEntry,
  type PakistanHolidayEmailAudience,
  type PakistanHolidayEmailEntry,
  type PakistanHolidayEmailStatus,
} from "@/lib/pakistan-holiday-emails";
import { normalizeEmail } from "@/lib/portal-users";
import { readPortalUsers } from "@/lib/portal-users-store";
import { requireEditor, requirePortalUser } from "@/lib/requireEditor";

function requirePakistanHolidaySender(
  request: Request,
): NextResponse | { email: string } {
  const email = normalizeEmail(request.headers.get("x-user-email") ?? "");
  if (!isPakistanHolidayEmailSender(email)) {
    return NextResponse.json(
      {
        error: `Only ${pakistanHolidayEmailSendersLabel()} can schedule or send Pakistan holiday emails.`,
      },
      { status: 403 },
    );
  }
  return { email };
}

function pakistanHolidaysFromStore(holidaysRaw: unknown) {
  const holidays = normalizeHolidays(holidaysRaw);
  const pakistan = holidays.countries.find(
    (country) => country.id === "pakistan",
  );
  return pakistan?.holidays ?? [];
}

function parseAudience(value: unknown): PakistanHolidayEmailAudience {
  return value === "client" ? "client" : "team";
}

export async function GET(request: Request) {
  const gate = await requirePortalUser(request);
  if (gate instanceof NextResponse) {
    return gate;
  }

  try {
    const [emailsRaw, holidaysRaw] = await Promise.all([
      readPakistanHolidayEmailsFile(),
      readHolidays(),
    ]);
    let store = normalizePakistanHolidayEmailsStore(emailsRaw);
    const promoted = promoteDueScheduledEntries(store);
    if (promoted.changed) {
      store = promoted.store;
      await writePakistanHolidayEmailsFile(store);
    }

    const pakistanHolidays = pakistanHolidaysFromStore(holidaysRaw);
    return NextResponse.json({
      entries: store.entries,
      statusByHolidayId: buildPakistanHolidayEmailStatusByHolidayId(
        pakistanHolidays,
        store,
        "team",
      ),
      clientStatusByHolidayId: buildPakistanHolidayEmailStatusByHolidayId(
        pakistanHolidays,
        store,
        "client",
      ),
      upcoming: buildPakistanHolidayEmailPreview(pakistanHolidays, store),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load Pakistan holiday email status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

type PostBody = {
  action?: "ensure" | "mark";
  audience?: PakistanHolidayEmailAudience;
  groupKey?: string;
  holidayDate?: string;
  holidayDates?: string[];
  holidayIds?: string[];
  holidayName?: string;
  status?: PakistanHolidayEmailStatus;
  scheduledFor?: string;
  sentAt?: string;
};

export async function POST(request: Request) {
  try {
    const body = (await request.json().catch(() => ({}))) as PostBody;
    const action = body.action === "mark" ? "mark" : "ensure";

    if (action === "mark") {
      const denied = await requireEditor(request);
      if (denied) {
        return denied;
      }
      const senderGate = requirePakistanHolidaySender(request);
      if (senderGate instanceof NextResponse) {
        return senderGate;
      }
      const holidayDate = String(body.holidayDate ?? "").trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(holidayDate)) {
        return NextResponse.json(
          { error: "holidayDate (YYYY-MM-DD) is required." },
          { status: 400 },
        );
      }
      const audience = parseAudience(body.audience);
      const holidayName = String(body.holidayName ?? "").trim();
      const holidayDates = Array.isArray(body.holidayDates)
        ? body.holidayDates
            .map((item) => String(item ?? "").trim())
            .filter((item) => /^\d{4}-\d{2}-\d{2}$/.test(item))
        : [];
      const holidayIds = Array.isArray(body.holidayIds)
        ? body.holidayIds
            .map((item) => String(item ?? "").trim())
            .filter(Boolean)
        : [];
      const dates = holidayDates.length > 0 ? holidayDates : [holidayDate];
      const groupKey = String(body.groupKey ?? "").trim();
      if (!groupKey) {
        return NextResponse.json(
          { error: "groupKey is required." },
          { status: 400 },
        );
      }
      const status: PakistanHolidayEmailStatus =
        body.status === "sent" ? "sent" : "scheduled";
      const store = normalizePakistanHolidayEmailsStore(
        await readPakistanHolidayEmailsFile(),
      );
      const entry: PakistanHolidayEmailEntry = {
        groupKey,
        audience,
        holidayDate: dates[0] ?? holidayDate,
        holidayDates: dates,
        holidayIds,
        holidayName,
        status,
        scheduledFor: String(body.scheduledFor ?? "").trim(),
        sentAt:
          status === "sent"
            ? String(body.sentAt ?? "").trim() || new Date().toISOString()
            : "",
      };
      const next = upsertPakistanHolidayEmailEntry(store, entry);
      await writePakistanHolidayEmailsFile(next);

      const pakistanHolidays = pakistanHolidaysFromStore(await readHolidays());
      return NextResponse.json({
        entries: next.entries,
        statusByHolidayId: buildPakistanHolidayEmailStatusByHolidayId(
          pakistanHolidays,
          next,
          "team",
        ),
        clientStatusByHolidayId: buildPakistanHolidayEmailStatusByHolidayId(
          pakistanHolidays,
          next,
          "client",
        ),
        upcoming: buildPakistanHolidayEmailPreview(pakistanHolidays, next),
        message:
          status === "sent" ? "Marked as sent." : "Marked as scheduled.",
      });
    }

    const denied = await requireEditor(request);
    if (denied) {
      return denied;
    }
    const senderGate = requirePakistanHolidaySender(request);
    if (senderGate instanceof NextResponse) {
      return senderGate;
    }

    const [holidaysRaw, templatesRaw, emailsRaw, users, profiles] =
      await Promise.all([
        readHolidays(),
        readEmailTemplates(),
        readPakistanHolidayEmailsFile(),
        readPortalUsers(),
        readClientProfiles(),
      ]);
    const pakistanHolidays = pakistanHolidaysFromStore(holidaysRaw);
    const templates = normalizeEmailTemplates(templatesRaw);
    let store = normalizePakistanHolidayEmailsStore(emailsRaw);
    const promoted = promoteDueScheduledEntries(store);
    if (promoted.changed) {
      store = promoted.store;
      await writePakistanHolidayEmailsFile(store);
    }

    const teamTemplate =
      templates.templates.find(
        (item) => item.id === PAKISTAN_HOLIDAY_EMAIL_TEMPLATE_ID,
      ) ?? null;
    const clientTemplate =
      templates.templates.find(
        (item) => item.id === PAKISTAN_HOLIDAY_CLIENT_EMAIL_TEMPLATE_ID,
      ) ?? null;
    if (!teamTemplate) {
      return NextResponse.json(
        { error: "Pakistan Holiday Notice template (template-16) is missing." },
        { status: 500 },
      );
    }
    if (!clientTemplate) {
      return NextResponse.json(
        {
          error:
            "Pakistan Holiday Notice (Clients) template (template-17) is missing.",
        },
        { status: 500 },
      );
    }

    const teamRecipients = users
      .map((user) => user.email)
      .filter((email) => Boolean(email));
    const clientRecipientGroups =
      collectPakistanHolidayClientRecipientGroups(profiles);
    const clientRecipients =
      collectPakistanHolidayClientRecipientEmails(profiles);

    const teamPlans = buildPakistanHolidayEmailPlans(
      pakistanHolidays,
      store,
      { subject: teamTemplate.subject, body: teamTemplate.body },
      { audience: "team" },
    );
    const clientPlans = buildPakistanHolidayEmailPlans(
      pakistanHolidays,
      store,
      { subject: clientTemplate.subject, body: clientTemplate.body },
      {
        audience: "client",
        recipientGroups: clientRecipientGroups,
      },
    );

    const planCount = teamPlans.length + clientPlans.length;
    let message =
      planCount === 0
        ? "No Pakistan holiday emails to schedule."
        : `${planCount} holiday email(s) ready to schedule.`;
    if (
      clientPlans.length === 0 &&
      clientRecipients.length === 0 &&
      teamPlans.length > 0
    ) {
      message +=
        " (No Client Profile emails found — client notices were skipped.)";
    } else if (
      clientRecipientGroups.length === 0 &&
      planCount === 0
    ) {
      message =
        "No Pakistan holiday emails to schedule. Add client emails in Client Profiles to enable client notices.";
    }

    return NextResponse.json({
      entries: store.entries,
      statusByHolidayId: buildPakistanHolidayEmailStatusByHolidayId(
        pakistanHolidays,
        store,
        "team",
      ),
      clientStatusByHolidayId: buildPakistanHolidayEmailStatusByHolidayId(
        pakistanHolidays,
        store,
        "client",
      ),
      upcoming: buildPakistanHolidayEmailPreview(pakistanHolidays, store),
      teamRecipients,
      clientRecipients,
      clientRecipientGroupCount: clientRecipientGroups.length,
      teamPlans,
      clientPlans,
      /** @deprecated use teamPlans / clientPlans */
      plans: teamPlans,
      recipients: teamRecipients,
      message,
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to process Pakistan holiday emails.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
