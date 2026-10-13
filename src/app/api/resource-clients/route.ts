import { NextResponse } from "next/server";
import {
  normalizeResourceClient,
  type ResourceClientEntry,
} from "@/lib/resource-clients";
import { readResourceClients, writeResourceClients } from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET() {
  try {
    const clients = await readResourceClients();
    return NextResponse.json({ clients });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load clients.";
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
      clients?: Partial<ResourceClientEntry>[];
    };

    if (!Array.isArray(body.clients)) {
      return NextResponse.json(
        { error: "Request body must include a clients array." },
        { status: 400 },
      );
    }

    const clients = body.clients
      .filter(
        (item): item is Partial<ResourceClientEntry> & { id: string } =>
          typeof item === "object" &&
          item !== null &&
          typeof item.id === "string",
      )
      .map(normalizeResourceClient);

    await writeResourceClients(clients);
    return NextResponse.json({ clients });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to save clients.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
