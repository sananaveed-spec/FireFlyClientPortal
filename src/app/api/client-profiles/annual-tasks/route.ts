import { NextResponse } from "next/server";
import {
  normalizeClientProfile,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import {
  mergeProtectedAnnualTasks,
  normalizeAnnualTasks,
  type AnnualTaskItem,
} from "@/lib/annual-tasks";
import { readClientProfiles, writeClientProfiles } from "@/lib/data-store";
import { seedAnnualTasksFromAcceptedProposal } from "@/lib/ignition/seed-annual-tasks";
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
      proposalSlug?: string;
      tasks?: AnnualTaskItem[];
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
      const slug =
        String(body.proposalSlug ?? "").trim() ||
        profile.activeProposalSlug.trim();
      if (slug) {
        profile = await seedAnnualTasksFromAcceptedProposal(profile, slug);
      }

      const next = [...profiles];
      next[index] = normalizeClientProfile(profile);
      await writeClientProfiles(next);
      return NextResponse.json({
        ok: true,
        profile: next[index],
        annualTasks: next[index]!.annualTasks,
      });
    }

    if (action === "save") {
      const gate = await requirePortalUser(request);
      if (gate instanceof NextResponse) {
        return gate;
      }
      const tasks = mergeProtectedAnnualTasks(
        normalizeAnnualTasks(profile.annualTasks),
        normalizeAnnualTasks(body.tasks),
      );
      const updated: ClientProfileEntry = normalizeClientProfile({
        ...profile,
        annualTasks: tasks,
      });
      const next = [...profiles];
      next[index] = updated;
      await writeClientProfiles(next);
      return NextResponse.json({
        ok: true,
        profile: updated,
        annualTasks: updated.annualTasks,
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
        : "Failed to update annual tasks.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
