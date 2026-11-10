import { NextResponse } from "next/server";
import {
  normalizeClientProfile,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import { readClientProfiles, writeClientProfiles } from "@/lib/data-store";
import { seedMonthlyTasksFromAcceptedProposal } from "@/lib/ignition/seed-monthly-tasks";
import {
  mergeProtectedMonthlyTasks,
  normalizeMonthlyTasks,
  type MonthlyTaskItem,
} from "@/lib/monthly-tasks";
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
      tasks?: MonthlyTaskItem[];
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
        profile = await seedMonthlyTasksFromAcceptedProposal(profile, slug);
      }

      const next = [...profiles];
      next[index] = normalizeClientProfile(profile);
      await writeClientProfiles(next);
      return NextResponse.json({
        ok: true,
        profile: next[index],
        monthlyTasks: next[index]!.monthlyTasks,
      });
    }

    if (action === "save") {
      const gate = await requirePortalUser(request);
      if (gate instanceof NextResponse) {
        return gate;
      }
      const tasks = mergeProtectedMonthlyTasks(
        normalizeMonthlyTasks(profile.monthlyTasks),
        normalizeMonthlyTasks(body.tasks),
      );
      const updated: ClientProfileEntry = normalizeClientProfile({
        ...profile,
        monthlyTasks: tasks,
      });
      const next = [...profiles];
      next[index] = updated;
      await writeClientProfiles(next);
      return NextResponse.json({
        ok: true,
        profile: updated,
        monthlyTasks: updated.monthlyTasks,
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
        : "Failed to update monthly tasks.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
