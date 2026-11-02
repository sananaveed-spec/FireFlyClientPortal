import { NextResponse } from "next/server";
import {
  clearBd8EmailAttachment,
  getBd8EmailAttachment,
  normalizeBd8EmailAttachment,
  normalizeBd8EmailAttachmentsStore,
  normalizeBdEmailAttachmentScheduleId,
  upsertBd8EmailAttachment,
  type BdEmailAttachmentScheduleId,
} from "@/lib/bd8-email-attachments";
import { currentMonthKey } from "@/lib/client-request-lists";
import {
  readBd8EmailAttachmentsFile,
  writeBd8EmailAttachmentsFile,
} from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

function scheduleIdFromRequest(
  url: URL,
  bodyScheduleId?: unknown,
): BdEmailAttachmentScheduleId {
  return normalizeBdEmailAttachmentScheduleId(
    bodyScheduleId ?? url.searchParams.get("scheduleId"),
  );
}

export async function GET(request: Request) {
  try {
    const url = new URL(request.url);
    const clientId = url.searchParams.get("clientId")?.trim() ?? "";
    const monthKey =
      url.searchParams.get("monthKey")?.trim() || currentMonthKey();
    const scheduleId = scheduleIdFromRequest(url);
    const parsed = await readBd8EmailAttachmentsFile();
    const store = normalizeBd8EmailAttachmentsStore(parsed, monthKey);

    // Persist prune so prior-month files are removed once.
    await writeBd8EmailAttachmentsFile(store);

    if (clientId) {
      return NextResponse.json({
        monthKey,
        scheduleId,
        attachment: getBd8EmailAttachment(
          store,
          clientId,
          monthKey,
          scheduleId,
        ),
      });
    }

    return NextResponse.json(store);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load BD8 email attachments.";
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
    const attachment = normalizeBd8EmailAttachment(body);
    if (!attachment) {
      return NextResponse.json(
        {
          error:
            "Request body must include clientId, monthKey, name, and contentBytes.",
        },
        { status: 400 },
      );
    }

    const monthKey = currentMonthKey();
    const scheduleId = attachment.scheduleId;
    const parsed = await readBd8EmailAttachmentsFile();
    const store = normalizeBd8EmailAttachmentsStore(parsed, monthKey);
    const next = upsertBd8EmailAttachment(store, {
      ...attachment,
      monthKey,
      scheduleId,
    });
    await writeBd8EmailAttachmentsFile(next);

    return NextResponse.json({
      monthKey,
      scheduleId,
      attachment: getBd8EmailAttachment(
        next,
        attachment.clientId,
        monthKey,
        scheduleId,
      ),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to save BD8 email attachment.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function DELETE(request: Request) {
  const denied = await requireEditor(request);
  if (denied) {
    return denied;
  }

  try {
    const url = new URL(request.url);
    const clientId = url.searchParams.get("clientId")?.trim() ?? "";
    const monthKey =
      url.searchParams.get("monthKey")?.trim() || currentMonthKey();
    const scheduleId = scheduleIdFromRequest(url);
    if (!clientId) {
      return NextResponse.json(
        { error: "clientId is required." },
        { status: 400 },
      );
    }

    const parsed = await readBd8EmailAttachmentsFile();
    const store = normalizeBd8EmailAttachmentsStore(parsed, monthKey);
    const next = clearBd8EmailAttachment(
      store,
      clientId,
      monthKey,
      scheduleId,
    );
    await writeBd8EmailAttachmentsFile(next);

    return NextResponse.json({
      monthKey,
      scheduleId,
      attachment: null,
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to remove BD8 email attachment.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
