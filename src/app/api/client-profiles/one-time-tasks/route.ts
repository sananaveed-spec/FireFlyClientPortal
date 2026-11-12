import { NextResponse } from "next/server";
import {
  normalizeClientProfile,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import { readClientProfiles, writeClientProfiles } from "@/lib/data-store";
import { seedOneTimeTasksFromAcceptedProposal } from "@/lib/ignition/seed-one-time-tasks";
import {
  hasQboSetupOneTimeTask,
  mergeProtectedOneTimeTasks,
  normalizeOneTimeTasks,
  type OneTimeTaskItem,
} from "@/lib/one-time-tasks";
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
      tasks?: OneTimeTaskItem[];
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

    let profile = profiles[index]!;

    if (action === "seed") {
      // Only seed from accepted proposal "Bookkeeping Services" heading.
      if (
        !hasQboSetupOneTimeTask(profile.oneTimeTasks) &&
        profile.activeProposalSlug.trim()
      ) {
        profile = await seedOneTimeTasksFromAcceptedProposal(
          profile,
          profile.activeProposalSlug,
        );
      }

      const next = [...profiles];
      next[index] = normalizeClientProfile(profile);
      await writeClientProfiles(next);
      return NextResponse.json({
        ok: true,
        profile: next[index],
        oneTimeTasks: next[index]!.oneTimeTasks,
      });
    }

    if (action === "save") {
      const gate = await requirePortalUser(request);
      if (gate instanceof NextResponse) {
        return gate;
      }
      const tasks = mergeProtectedOneTimeTasks(
        normalizeOneTimeTasks(profile.oneTimeTasks),
        normalizeOneTimeTasks(body.tasks),
      );
      const updated: ClientProfileEntry = normalizeClientProfile({
        ...profile,
        oneTimeTasks: tasks,
      });
      const next = [...profiles];
      next[index] = updated;
      await writeClientProfiles(next);
      return NextResponse.json({
        ok: true,
        profile: updated,
        oneTimeTasks: updated.oneTimeTasks,
      });
    }

    return NextResponse.json(
      { error: 'Expected action "seed" or "save".' },
      { status: 400 },
    );
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to update one-time tasks.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
