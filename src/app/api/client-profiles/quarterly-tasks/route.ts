import { NextResponse } from "next/server";
import {
  normalizeClientProfile,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import { readClientProfiles, writeClientProfiles } from "@/lib/data-store";
import {
  mergeProtectedQuarterlyTasks,
  normalizeQuarterlyTasks,
  type QuarterlyTaskItem,
} from "@/lib/quarterly-tasks";
import { requireEditor, requirePortalUser } from "@/lib/requireEditor";

export async function POST(request: Request) {
  const denied = await requireEditor(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json().catch(() => ({}))) as {
      action?: string;
      profileId?: string;
      tasks?: QuarterlyTaskItem[];
    };
    const action = String(body.action ?? "").trim().toLowerCase();
    const profileId = String(body.profileId ?? "").trim();
    if (!profileId) {
      return NextResponse.json(
        { error: "profileId is required." },
        { status: 400 },
      );
    }

    const profiles = await readClientProfiles();
    const index = profiles.findIndex((profile) => profile.id === profileId);
    if (index < 0) {
      return NextResponse.json(
        { error: "Client profile not found." },
        { status: 404 },
      );
    }

    const profile = profiles[index]!;

    if (action === "save") {
      const gate = await requirePortalUser(request);
      if (gate instanceof NextResponse) {
        return gate;
      }
      const tasks = mergeProtectedQuarterlyTasks(
        normalizeQuarterlyTasks(profile.quarterlyTasks),
        normalizeQuarterlyTasks(body.tasks),
      );
      const updated: ClientProfileEntry = normalizeClientProfile({
        ...profile,
        quarterlyTasks: tasks,
      });
      const next = [...profiles];
      next[index] = updated;
      await writeClientProfiles(next);
      return NextResponse.json({
        ok: true,
        profile: updated,
        quarterlyTasks: updated.quarterlyTasks,
      });
    }

    return NextResponse.json(
      { error: 'Expected action "save".' },
      { status: 400 },
    );
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to update quarterly tasks.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
