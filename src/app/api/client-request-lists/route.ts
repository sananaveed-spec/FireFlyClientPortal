import { NextResponse } from "next/server";
import {
  normalizeClientRequestList,
  normalizeClientRequestListsStore,
  promoteAllListsForNewMonth,
  upsertClientRequestList,
} from "@/lib/client-request-lists";
import {
  readClientRequestListsFile,
  writeClientRequestListsFile,
} from "@/lib/data-store";
import { requirePortalUser } from "@/lib/requireEditor";

export async function GET() {
  try {
    const parsed = await readClientRequestListsFile();
    const current = normalizeClientRequestListsStore(parsed);
    const promoted = promoteAllListsForNewMonth(current);
    if (promoted.changed) {
      await writeClientRequestListsFile(promoted.store);
    }
    return NextResponse.json(promoted.store);
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
    const body = await request.json();
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

    const current = normalizeClientRequestListsStore(
      await readClientRequestListsFile(),
    );
    const withPromotion = promoteAllListsForNewMonth(current).store;
    const next = upsertClientRequestList(withPromotion, list);
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
