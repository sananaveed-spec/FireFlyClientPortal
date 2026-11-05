import { NextResponse } from "next/server";
import { currentMonthKey } from "@/lib/client-request-lists";
import {
  addEpsFridayAttachment,
  clearEpsFridayAttachments,
  getEpsFridayAttachments,
  isAllowedEpsFridaySpreadsheetName,
  normalizeEpsFridayAttachment,
  normalizeEpsFridayAttachmentsStore,
  removeEpsFridayAttachment,
} from "@/lib/eps-friday-attachments";
import { isEpsFridayWeekScheduleId } from "@/lib/eps-friday-weeks";
import {
  readEpsFridayAttachmentsFile,
  writeEpsFridayAttachmentsFile,
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
    const store = normalizeEpsFridayAttachmentsStore(
      await readEpsFridayAttachmentsFile(),
      monthKey,
    );
    await writeEpsFridayAttachmentsFile(store);

    if (clientId && scheduleId) {
      return NextResponse.json({
        monthKey,
        scheduleId,
        attachments: getEpsFridayAttachments(
          store,
          clientId,
          scheduleId,
          monthKey,
        ),
      });
    }

    return NextResponse.json(store);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load EPS Friday attachments.";
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
    const files = Array.isArray(body?.attachments)
      ? body.attachments
      : [body];
    const monthKey = currentMonthKey();
    let store = normalizeEpsFridayAttachmentsStore(
      await readEpsFridayAttachmentsFile(),
      monthKey,
    );

    let clientId = "";
    let scheduleId = "";
    for (const raw of files) {
      const attachment = normalizeEpsFridayAttachment({
        ...raw,
        monthKey: String(raw?.monthKey ?? monthKey),
      });
      if (!attachment) {
        return NextResponse.json(
          {
            error:
              "Each file needs clientId, scheduleId (eps-week-N-friday), name, and contentBytes.",
          },
          { status: 400 },
        );
      }
      if (!isAllowedEpsFridaySpreadsheetName(attachment.name)) {
        return NextResponse.json(
          { error: `Only .xlsx, .xls, or .csv allowed (${attachment.name}).` },
          { status: 400 },
        );
      }
      clientId = attachment.clientId;
      scheduleId = attachment.scheduleId;
      store = addEpsFridayAttachment(store, {
        ...attachment,
        monthKey,
      });
    }

    await writeEpsFridayAttachmentsFile(store);
    return NextResponse.json({
      monthKey,
      scheduleId,
      attachments: getEpsFridayAttachments(
        store,
        clientId,
        scheduleId,
        monthKey,
      ),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to save EPS Friday attachments.";
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
    const scheduleId = url.searchParams.get("scheduleId")?.trim() ?? "";
    const attachmentId = url.searchParams.get("attachmentId")?.trim() ?? "";
    const monthKey =
      url.searchParams.get("monthKey")?.trim() || currentMonthKey();

    if (!clientId || !isEpsFridayWeekScheduleId(scheduleId)) {
      return NextResponse.json(
        { error: "clientId and eps-week scheduleId are required." },
        { status: 400 },
      );
    }

    let store = normalizeEpsFridayAttachmentsStore(
      await readEpsFridayAttachmentsFile(),
      monthKey,
    );
    store = attachmentId
      ? removeEpsFridayAttachment(
          store,
          clientId,
          scheduleId,
          attachmentId,
          monthKey,
        )
      : clearEpsFridayAttachments(store, clientId, scheduleId, monthKey);
    await writeEpsFridayAttachmentsFile(store);

    return NextResponse.json({
      monthKey,
      scheduleId,
      attachments: getEpsFridayAttachments(
        store,
        clientId,
        scheduleId,
        monthKey,
      ),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to clear EPS Friday attachments.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
