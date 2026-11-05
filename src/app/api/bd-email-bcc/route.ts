import { NextResponse } from "next/server";
import {
  getBdEmailBcc,
  normalizeBdEmailBccStore,
  upsertBdEmailBcc,
} from "@/lib/bd-email-bcc";
import { readBdEmailBccFile, writeBdEmailBccFile } from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export const runtime = "nodejs";

export async function GET(request: Request) {
  try {
    const url = new URL(request.url);
    const clientId = url.searchParams.get("clientId")?.trim() ?? "";
    const scheduleId = url.searchParams.get("scheduleId")?.trim() ?? "";
    const store = normalizeBdEmailBccStore(await readBdEmailBccFile());

    if (clientId && scheduleId) {
      return NextResponse.json({
        clientId,
        scheduleId,
        bcc: getBdEmailBcc(store, clientId, scheduleId),
      });
    }

    return NextResponse.json(store);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load BD email BCC.";
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
      clientId?: string;
      scheduleId?: string;
      bcc?: string;
    };
    const clientId = String(body.clientId ?? "").trim();
    const scheduleId = String(body.scheduleId ?? "").trim();
    if (!clientId || !scheduleId) {
      return NextResponse.json(
        { error: "clientId and scheduleId are required." },
        { status: 400 },
      );
    }

    const current = normalizeBdEmailBccStore(await readBdEmailBccFile());
    const next = upsertBdEmailBcc(current, {
      clientId,
      scheduleId,
      bcc: String(body.bcc ?? ""),
    });
    await writeBdEmailBccFile(next);
    return NextResponse.json({
      clientId,
      scheduleId,
      bcc: getBdEmailBcc(next, clientId, scheduleId),
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to save BD email BCC.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
