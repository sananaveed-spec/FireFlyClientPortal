import { NextResponse } from "next/server";
import { normalizeHolidays } from "@/lib/holidays";
import { readHolidays, writeHolidays } from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET() {
  try {
    const data = await readHolidays();
    return NextResponse.json(data);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load holidays.";
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
    const data = normalizeHolidays(body);
    await writeHolidays(data);
    return NextResponse.json(data);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to save holidays.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
