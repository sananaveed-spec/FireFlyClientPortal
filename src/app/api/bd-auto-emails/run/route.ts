import { NextResponse } from "next/server";
import {
  isBdAutoEmailCronAuthorized,
  runBdAutoEmails,
} from "@/lib/bd-auto-emails";
import { isBdAutoSendScheduleId } from "@/lib/bd-auto-email-status";

export const runtime = "nodejs";

/**
 * Cron entry: send due BD1 / BD7 / BD8 emails as supervisor (app-only Graph).
 *
 * Auth: Authorization: Bearer $BD_AUTO_EMAIL_CRON_SECRET
 * Optional: ?force=1 to re-send even if lastSentAt is set.
 * Optional: ?scheduleId=bd1 to limit to one BD.
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

    const result = await runBdAutoEmails({ force, scheduleIds });
    return NextResponse.json(result, { status: result.ok ? 200 : 207 });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "BD auto-email run failed.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
