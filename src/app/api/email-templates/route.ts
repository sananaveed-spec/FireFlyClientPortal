import { NextResponse } from "next/server";
import { normalizeEmailTemplates } from "@/lib/email-templates";
import { readEmailTemplates, writeEmailTemplates } from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET() {
  try {
    const data = await readEmailTemplates();
    return NextResponse.json(data);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load email templates.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  const denied = await requireEditor(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json()) as unknown;
    const data = normalizeEmailTemplates(body);
    await writeEmailTemplates(data);
    return NextResponse.json(data);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to save email templates.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
