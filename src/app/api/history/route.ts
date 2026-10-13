import { NextResponse } from "next/server";
import type { HistoryInput } from "@/lib/history";
import { appendHistory, readHistory, writeHistory } from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET() {
  try {
    const entries = await readHistory();
    return NextResponse.json({ entries });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load history.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  const denied = await requireEditor(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json()) as HistoryInput;
    const entry = await appendHistory(body);
    return NextResponse.json({ entry });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to save history.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function DELETE(request: Request) {
  const denied = await requireEditor(request);
  if (denied) {
    return denied;
  }

  try {
    await writeHistory([]);
    return NextResponse.json({ entries: [] });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to clear history.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
