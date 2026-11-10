import type { ClientProfileEntry } from "@/lib/client-profiles";
import { seedAnnualTasksFromAcceptedProposal } from "@/lib/ignition/seed-annual-tasks";
import { seedMonthlyTasksFromAcceptedProposal } from "@/lib/ignition/seed-monthly-tasks";
import { seedOneTimeTasksFromAcceptedProposal } from "@/lib/ignition/seed-one-time-tasks";

/** Seed One-Time + Annual + Monthly tasks from the active proposal (empty if none). */
export async function seedClientEngagementTasksFromProposal(
  profile: ClientProfileEntry,
  proposalSlug?: string,
): Promise<ClientProfileEntry> {
  let next = await seedOneTimeTasksFromAcceptedProposal(profile, proposalSlug);
  next = await seedAnnualTasksFromAcceptedProposal(next, proposalSlug);
  next = await seedMonthlyTasksFromAcceptedProposal(next, proposalSlug);
  return next;
}
