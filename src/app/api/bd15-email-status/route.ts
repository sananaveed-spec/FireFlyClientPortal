import { NextResponse } from "next/server";
import {
  addBd15EmailStatusComment,
  getBd15EmailStatus,
  markBd15EmailSent,
  normalizeBd15EmailStatusStore,
} from "@/lib/bd15-email-status";
import { currentMonthKey } from "@/lib/client-request-lists";
import {
  readBd15EmailStatusFile,
  writeBd15EmailStatusFile,
} from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET(request: Request) {
  try {
    const url = new URL(request.url);
    const clientId = url.searchParams.get("clientId")?.trim() ?? "";
    const monthKey =
      url.searchParams.get("monthKey")?.trim() || currentMonthKey();
    const parsed = await readBd15EmailStatusFile();
    const store = normalizeBd15EmailStatusStore(parsed, monthKey);
    await writeBd15EmailStatusFile(store);

    if (clientId) {
      return NextResponse.json({
        monthKey,
        status: getBd15EmailStatus(store, clientId, monthKey),
      });
    }

    return NextResponse.json(store);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load BD15 email status.";
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
      action?: string;
      clientId?: string;
      monthKey?: string;
      text?: string;
      sentAt?: string;
    };
    const clientId = String(body.clientId ?? "").trim();
    const monthKey = String(body.monthKey ?? "").trim() || currentMonthKey();
    const action = String(body.action ?? "").trim();
    if (!clientId || !action) {
      return NextResponse.json(
        { error: "clientId and action are required." },
        { status: 400 },
      );
    }

    const parsed = await readBd15EmailStatusFile();
    const store = normalizeBd15EmailStatusStore(parsed, monthKey);
    let next = store;

    if (action === "mark-sent") {
      next = markBd15EmailSent(
        store,
        clientId,
        monthKey,
        typeof body.sentAt === "string" && body.sentAt.trim()
          ? body.sentAt.trim()
          : new Date().toISOString(),
      );
    } else if (action === "add-comment") {
      const text = String(body.text ?? "").trim();
      if (!text) {
        return NextResponse.json(
          { error: "Comment text is required." },
          { status: 400 },
        );
      }
      next = addBd15EmailStatusComment(store, clientId, text, monthKey);
    } else {
      return NextResponse.json(
        { error: 'action must be "mark-sent" or "add-comment".' },
        { status: 400 },
      );
    }

    await writeBd15EmailStatusFile(next);
    return NextResponse.json({
      monthKey,
      status: getBd15EmailStatus(next, clientId, monthKey),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to update BD15 email status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
