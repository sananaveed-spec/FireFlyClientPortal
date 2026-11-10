import { NextResponse } from "next/server";
import {
  readHolidays,
  readPakistanHolidayEmailsFile,
  writePakistanHolidayEmailsFile,
} from "@/lib/data-store";
import { normalizeHolidays } from "@/lib/holidays";
import { runPakistanHolidayAutoEmails } from "@/lib/pakistan-holiday-auto-emails";
import {
  buildPakistanHolidayEmailPreview,
  buildPakistanHolidayEmailStatusByHolidayId,
  isPakistanHolidayEmailSender,
  normalizePakistanHolidayEmailsStore,
  pakistanHolidayEmailSendersLabel,
  upsertPakistanHolidayEmailEntry,
  type PakistanHolidayEmailAudience,
  type PakistanHolidayEmailEntry,
  type PakistanHolidayEmailStatus,
} from "@/lib/pakistan-holiday-emails";
import { normalizeEmail } from "@/lib/portal-users";
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
    const store = normalizePakistanHolidayEmailsStore(emailsRaw);
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
  action?: "ensure" | "mark" | "auto-run";
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
    const action = body.action === "mark" ? "mark" : "auto-run";

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

    // ensure / auto-run: queue upcoming + send due via app-only Mail.Send (BD cron).
    const denied = await requireEditor(request);
    if (denied) {
      return denied;
    }
    const senderGate = requirePakistanHolidaySender(request);
    if (senderGate instanceof NextResponse) {
      return senderGate;
    }

    const run = await runPakistanHolidayAutoEmails();
    const [holidaysRaw, emailsRaw] = await Promise.all([
      readHolidays(),
      readPakistanHolidayEmailsFile(),
    ]);
    const pakistanHolidays = pakistanHolidaysFromStore(holidaysRaw);
    const store = normalizePakistanHolidayEmailsStore(emailsRaw);
    const parts: string[] = [];
    if (run.queued > 0) {
      parts.push(`${run.queued} queued`);
    }
    if (run.sent.length > 0) {
      parts.push(`${run.sent.length} sent`);
    }
    if (run.errors.length > 0) {
      parts.push(`${run.errors.length} failed`);
    }
    const message =
      parts.length > 0
        ? `Holiday emails (auto): ${parts.join(", ")}.`
        : run.skipped[0]?.reason ||
          "No Pakistan holiday emails due right now (server cron sends when due).";

    return NextResponse.json({
      ok: run.ok,
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
      autoRun: run,
      teamPlans: [],
      clientPlans: [],
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
