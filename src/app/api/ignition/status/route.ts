import { NextResponse } from "next/server";
import { readIgnitionClientsStore } from "@/lib/ignition/clients";
import {
  ignitionConnectionStatus,
  readIgnitionTokens,
} from "@/lib/ignition/tokens";
import { requireAdmin } from "@/lib/requireEditor";

export async function GET(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const tokens = await readIgnitionTokens();
    const status = ignitionConnectionStatus(tokens);
    const store = await readIgnitionClientsStore();

    return NextResponse.json({
      connected: status.connected,
      configured: status.configured,
      expiresAt: status.expiresAt,
      syncedAt: store.syncedAt || null,
      clientCount: store.clients.length,
      clients: store.clients.map((client) => ({
        id: client.id,
        name: client.name,
        email: client.email,
        status: client.status,
      })),
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load Ignition status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
