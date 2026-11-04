import { NextResponse } from "next/server";
import {
  getBdAutoEmailStatus,
  markBdAutoEmailSent,
  normalizeBdAutoEmailStatusStore,
  isBdAutoSendScheduleId,
} from "@/lib/bd-auto-email-status";
import { currentMonthKey } from "@/lib/client-request-lists";
import {
  readBdAutoEmailStatusFile,
  writeBdAutoEmailStatusFile,
} from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export const runtime = "nodejs";

export async function GET(request: Request) {
  try {
    const url = new URL(request.url);
    const clientId = url.searchParams.get("clientId")?.trim() ?? "";
    const scheduleId = url.searchParams.get("scheduleId")?.trim() ?? "";
    const monthKey =
      url.searchParams.get("monthKey")?.trim() || currentMonthKey();
    const parsed = await readBdAutoEmailStatusFile();
    const store = normalizeBdAutoEmailStatusStore(parsed, monthKey);

    if (clientId && scheduleId) {
      return NextResponse.json({
        monthKey,
        status: getBdAutoEmailStatus(store, clientId, scheduleId, monthKey),
      });
    }

    if (clientId) {
      return NextResponse.json({
        monthKey,
        statuses: store.statuses.filter((item) => item.clientId === clientId),
      });
    }

    return NextResponse.json({ monthKey, ...store });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load BD auto-email status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

/** Mark sent after supervisor manual Schedule (prevents cron double-send). */
export async function PUT(request: Request) {
  const denied = await requireEditor(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json()) as {
      action?: string;
      clientId?: string;
      scheduleId?: string;
      monthKey?: string;
      lastSentAt?: string;
      fromEmail?: string;
      source?: "manual-schedule" | "manual-mark";
    };
    const clientId = String(body.clientId ?? "").trim();
    const scheduleId = String(body.scheduleId ?? "").trim();
    const monthKey = String(body.monthKey ?? "").trim() || currentMonthKey();
    const action = String(body.action ?? "").trim();

    if (!clientId || !scheduleId || !action) {
      return NextResponse.json(
        { error: "clientId, scheduleId, and action are required." },
        { status: 400 },
      );
    }
    if (!isBdAutoSendScheduleId(scheduleId)) {
      return NextResponse.json(
        { error: "scheduleId must be bd1, bd7, or bd8." },
        { status: 400 },
      );
    }

    const parsed = await readBdAutoEmailStatusFile();
    let store = normalizeBdAutoEmailStatusStore(parsed, monthKey);

    if (action === "mark-sent") {
      store = markBdAutoEmailSent(store, {
        clientId,
        scheduleId,
        monthKey,
        lastSentAt: body.lastSentAt,
        fromEmail: body.fromEmail,
        source: body.source === "manual-mark" ? "manual-mark" : "manual-schedule",
      });
    } else {
      return NextResponse.json(
        { error: 'Unknown action. Use "mark-sent".' },
        { status: 400 },
      );
    }

    await writeBdAutoEmailStatusFile(store);
    return NextResponse.json({
      monthKey,
      status: getBdAutoEmailStatus(store, clientId, scheduleId, monthKey),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to update BD auto-email status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
