import type { ClientProfileEntry } from "@/lib/client-profiles";
import { seedAnnualTasksFromAcceptedProposal } from "@/lib/ignition/seed-annual-tasks";

/**
 * On proposal accept: seed Annual Tasks only.
 * One-Time + Monthly Bookkeeping tasks are seeded when Resource Allocation
 * is created/saved (see Resource Allocation panel + task seed APIs).
 */
export async function seedClientEngagementTasksFromProposal(
  profile: ClientProfileEntry,
  proposalSlug?: string,
): Promise<ClientProfileEntry> {
  return seedAnnualTasksFromAcceptedProposal(profile, proposalSlug);
}
