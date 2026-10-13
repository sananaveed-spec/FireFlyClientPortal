import { NextResponse } from "next/server";
import { normalizeLookupData } from "@/lib/lookup-data";
import { readLookupData, writeLookupData } from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET() {
  try {
    const data = await readLookupData();
    return NextResponse.json(data);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load lookup data.";
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
    const data = normalizeLookupData(body);
    await writeLookupData(data);
    return NextResponse.json(data);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to save lookup data.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
