import { NextResponse } from "next/server";
import {
  normalizeClientScheduleSettings,
  normalizeClientScheduleSettingsStore,
  upsertClientScheduleSettings,
} from "@/lib/client-schedule-settings";
import {
  readClientScheduleSettingsFile,
  writeClientScheduleSettingsFile,
} from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET() {
  try {
    const parsed = await readClientScheduleSettingsFile();
    const store = normalizeClientScheduleSettingsStore(parsed);
    return NextResponse.json(store);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load schedule settings.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  const denied = await requireEditor(request);
  if (denied) {
    return denied;
  }

  try {
    const body = await request.json();
    const entry = normalizeClientScheduleSettings(body);
    if (!entry) {
      return NextResponse.json(
        {
          error:
            "Request body must include clientId and activeScheduleIds array.",
        },
        { status: 400 },
      );
    }

    const current = normalizeClientScheduleSettingsStore(
      await readClientScheduleSettingsFile(),
    );
    const next = upsertClientScheduleSettings(current, entry);
    await writeClientScheduleSettingsFile(next);
    return NextResponse.json(next);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to save schedule settings.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
