import { NextResponse } from "next/server";
import {
  currentMonthKey,
  normalizeClientRequestList,
  normalizeClientRequestListsStore,
  parseMonthKeyToClock,
  promoteAllListsForNewMonth,
  upsertClientRequestList,
} from "@/lib/client-request-lists";
import {
  readClientRequestListsFile,
  writeClientRequestListsFile,
} from "@/lib/data-store";
import { requirePortalUser } from "@/lib/requireEditor";

function clockFromRequest(request: Request): Date {
  const url = new URL(request.url);
  const monthKey =
    url.searchParams.get("monthKey")?.trim() ||
    url.searchParams.get("asOf")?.trim() ||
    "";
  return parseMonthKeyToClock(monthKey) ?? new Date();
}

export async function GET(request: Request) {
  try {
    const clock = clockFromRequest(request);
    const parsed = await readClientRequestListsFile();
    const current = normalizeClientRequestListsStore(parsed);
    const promoted = promoteAllListsForNewMonth(current, clock);
    if (promoted.changed) {
      await writeClientRequestListsFile(promoted.store);
    }
    return NextResponse.json({
      ...promoted.store,
      monthKey: currentMonthKey(clock),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load request lists.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  const gate = await requirePortalUser(request);
  if (gate instanceof NextResponse) {
    return gate;
  }

  try {
    const body = (await request.json()) as Record<string, unknown>;
    const list = normalizeClientRequestList(body);
    if (!list) {
      return NextResponse.json(
        {
          error:
            "Request body must include clientId, scheduleId, requestItems, and actionItems.",
        },
        { status: 400 },
      );
    }

    const bodyMonth =
      typeof body.monthKey === "string" ? body.monthKey.trim() : "";
    const clock = parseMonthKeyToClock(bodyMonth) ?? clockFromRequest(request);
    const current = normalizeClientRequestListsStore(
      await readClientRequestListsFile(),
    );
    const withPromotion = promoteAllListsForNewMonth(current, clock).store;
    const next = upsertClientRequestList(withPromotion, list, clock);
    await writeClientRequestListsFile(next);
    return NextResponse.json(next);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to save request list.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
