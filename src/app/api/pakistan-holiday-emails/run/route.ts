import { NextResponse } from "next/server";
import { isBdAutoEmailCronAuthorized } from "@/lib/bd-auto-emails";
import { runPakistanHolidayAutoEmails } from "@/lib/pakistan-holiday-auto-emails";

export const runtime = "nodejs";

/**
 * Dedicated cron for Pakistan holiday emails only.
 * Team: 9:00 AM Pakistan time · Clients: 9:00 AM each client's time zone.
 *
 * Auth: Authorization: Bearer $BD_AUTO_EMAIL_CRON_SECRET
 * (same secret as BD auto-emails)
 *
 * Recommended: hourly so every client timezone's 9:00 AM is caught soon after.
 * Optional: ?force=1 to ignore the post-9am send window.
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
    const result = await runPakistanHolidayAutoEmails({ force });
    return NextResponse.json(result, { status: result.ok ? 200 : 207 });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Pakistan holiday auto-email run failed.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
