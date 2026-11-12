import type { ClientProfileEntry } from "@/lib/client-profiles";
import { seedAnnualTasksFromAcceptedProposal } from "@/lib/ignition/seed-annual-tasks";

/**
 * On proposal accept: seed Annual Tasks from Income Tax Compliance Service(s)
 * heading only. One-Time + Monthly tasks seed on Resource Allocation save.
 */
export async function seedClientEngagementTasksFromProposal(
  profile: ClientProfileEntry,
  proposalSlug?: string,
): Promise<ClientProfileEntry> {
  return seedAnnualTasksFromAcceptedProposal(profile, proposalSlug);
}
