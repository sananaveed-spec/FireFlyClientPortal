import { NextResponse } from "next/server";
import {
  normalizeClientAdjustableSettings,
  normalizeFlexibleAdjustableSettingsStore,
  upsertClientAdjustableRows,
} from "@/lib/flexible-adjustable-settings";
import {
  readFlexibleAdjustableSettingsFile,
  writeFlexibleAdjustableSettingsFile,
} from "@/lib/data-store";
import { requirePortalUser } from "@/lib/requireEditor";

export async function GET() {
  try {
    const parsed = await readFlexibleAdjustableSettingsFile();
    const store = normalizeFlexibleAdjustableSettingsStore(parsed);
    return NextResponse.json(store);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load adjustable settings.";
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
    const client = normalizeClientAdjustableSettings(body);
    if (!client) {
      return NextResponse.json(
        { error: "Request body must include clientId and rows." },
        { status: 400 },
      );
    }

    const current = normalizeFlexibleAdjustableSettingsStore(
      await readFlexibleAdjustableSettingsFile(),
    );
    const next = upsertClientAdjustableRows(
      current,
      client.clientId,
      client.rows,
    );
    await writeFlexibleAdjustableSettingsFile(next);
    return NextResponse.json(next);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to save adjustable settings.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
