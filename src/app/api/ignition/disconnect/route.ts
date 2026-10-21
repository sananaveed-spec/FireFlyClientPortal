import { NextResponse } from "next/server";
import { clearIgnitionTokens } from "@/lib/ignition/tokens";
import { requireAdmin } from "@/lib/requireEditor";

export async function POST(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    await clearIgnitionTokens();
    return NextResponse.json({ connected: false });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to disconnect Ignition.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
