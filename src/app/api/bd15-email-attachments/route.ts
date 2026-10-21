import { NextResponse } from "next/server";
import {
  addBd15EmailAttachment,
  getBd15EmailAttachments,
  normalizeBd15EmailAttachment,
  normalizeBd15EmailAttachmentsStore,
  removeBd15EmailAttachment,
} from "@/lib/bd15-email-attachments";
import { currentMonthKey } from "@/lib/client-request-lists";
import {
  readBd15EmailAttachmentsFile,
  writeBd15EmailAttachmentsFile,
} from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET(request: Request) {
  try {
    const url = new URL(request.url);
    const clientId = url.searchParams.get("clientId")?.trim() ?? "";
    const monthKey =
      url.searchParams.get("monthKey")?.trim() || currentMonthKey();
    const parsed = await readBd15EmailAttachmentsFile();
    const store = normalizeBd15EmailAttachmentsStore(parsed, monthKey);

    // Persist prune so prior-month files are removed once.
    await writeBd15EmailAttachmentsFile(store);

    if (clientId) {
      return NextResponse.json({
        monthKey,
        attachments: getBd15EmailAttachments(store, clientId, monthKey),
      });
    }

    return NextResponse.json(store);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load BD15 email attachments.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  const denied = await requireEditor(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json()) as {
      replaceSameName?: boolean;
      [key: string]: unknown;
    };
    const attachment = normalizeBd15EmailAttachment(body);
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
    const parsed = await readBd15EmailAttachmentsFile();
    const store = normalizeBd15EmailAttachmentsStore(parsed, monthKey);
    const next = addBd15EmailAttachment(
      store,
      {
        ...attachment,
        monthKey,
      },
      { replaceSameName: Boolean(body.replaceSameName) },
    );
    await writeBd15EmailAttachmentsFile(next);

    return NextResponse.json({
      monthKey,
      attachments: getBd15EmailAttachments(next, attachment.clientId, monthKey),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to save BD15 email attachment.";
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
    const attachmentId = url.searchParams.get("id")?.trim() ?? "";
    const monthKey =
      url.searchParams.get("monthKey")?.trim() || currentMonthKey();
    if (!clientId || !attachmentId) {
      return NextResponse.json(
        { error: "clientId and id are required." },
        { status: 400 },
      );
    }

    const parsed = await readBd15EmailAttachmentsFile();
    const store = normalizeBd15EmailAttachmentsStore(parsed, monthKey);
    const next = removeBd15EmailAttachment(
      store,
      clientId,
      attachmentId,
      monthKey,
    );
    await writeBd15EmailAttachmentsFile(next);

    return NextResponse.json({
      monthKey,
      attachments: getBd15EmailAttachments(next, clientId, monthKey),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to remove BD15 email attachment.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
