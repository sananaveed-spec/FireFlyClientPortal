import { NextResponse } from "next/server";
import {
  isBdAutoEmailCronAuthorized,
  runBdAutoEmails,
} from "@/lib/bd-auto-emails";
import { isBdAutoSendScheduleId } from "@/lib/bd-auto-email-status";
import { runPakistanHolidayAutoEmails } from "@/lib/pakistan-holiday-auto-emails";

export const runtime = "nodejs";

/**
 * Cron entry: send due BD emails as supervisor (app-only Graph), and also
 * queue/send Pakistan holiday team + client notices (same secret / schedule).
 *
 * Auth: Authorization: Bearer $BD_AUTO_EMAIL_CRON_SECRET
 * Optional: ?force=1 to re-send even if lastSentAt is set (still requires send window).
 * Optional: ?scheduleId=bd1 to limit to one BD (holiday auto-send still runs).
 */
export async function POST(request: Request) {
  try {
    if (!isBdAutoEmailCronAuthorized(request)) {
      return NextResponse.json(
        {
          error:
            "Unauthorized. Set BD_AUTO_EMAIL_CRON_SECRET and pass Bearer token.",
        },
        { status: 401 },
      );
    }

    const url = new URL(request.url);
    const force = url.searchParams.get("force") === "1";
    const scheduleIdParam = url.searchParams.get("scheduleId")?.trim() ?? "";
    const scheduleIds = isBdAutoSendScheduleId(scheduleIdParam)
      ? [scheduleIdParam]
      : undefined;

    const [result, pakistanHolidayEmails] = await Promise.all([
      runBdAutoEmails({ force, scheduleIds }),
      runPakistanHolidayAutoEmails({ force }),
    ]);
    const ok = result.ok && pakistanHolidayEmails.ok;
    return NextResponse.json(
      { ...result, pakistanHolidayEmails },
      { status: ok ? 200 : 207 },
    );
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "BD auto-email run failed.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
