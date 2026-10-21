import { NextResponse } from "next/server";
import {
  readIgnitionClientsStore,
  syncIgnitionClients,
} from "@/lib/ignition/clients";
import { requireAdmin } from "@/lib/requireEditor";

export async function GET(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const store = await readIgnitionClientsStore();
    return NextResponse.json({
      syncedAt: store.syncedAt || null,
      clients: store.clients.map((client) => ({
        id: client.id,
        name: client.name,
        email: client.email,
        status: client.status,
      })),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load synced Ignition clients.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const store = await syncIgnitionClients();
    return NextResponse.json({
      syncedAt: store.syncedAt,
      clients: store.clients.map((client) => ({
        id: client.id,
        name: client.name,
        email: client.email,
        status: client.status,
      })),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to sync Ignition clients.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
