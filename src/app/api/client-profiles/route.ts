import { NextResponse } from "next/server";
import {
  normalizeClientProfile,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import { readClientProfiles, writeClientProfiles } from "@/lib/data-store";
import { requireEditor } from "@/lib/requireEditor";

export async function GET() {
  try {
    const profiles = await readClientProfiles();
    return NextResponse.json({ profiles });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load profiles.";
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
      profiles?: Partial<ClientProfileEntry>[];
    };

    if (!Array.isArray(body.profiles)) {
      return NextResponse.json(
        { error: "Request body must include a profiles array." },
        { status: 400 },
      );
    }

    const profiles = body.profiles
      .filter(
        (item): item is Partial<ClientProfileEntry> & { id: string } =>
          typeof item === "object" &&
          item !== null &&
          typeof item.id === "string",
      )
      .map(normalizeClientProfile);

    await writeClientProfiles(profiles);
    return NextResponse.json({ profiles });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to save profiles.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
