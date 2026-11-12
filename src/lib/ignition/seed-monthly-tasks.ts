import {
  ensureBookkeepingMonthlyTasks,
} from "@/lib/monthly-tasks";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import { fetchIgnitionProposalView } from "@/lib/ignition/proposals";
import { proposalHasBookkeepingServicesHeading } from "@/lib/one-time-tasks";

/**
 * Called when Resource Allocation is saved (not on proposal accept).
 * If the client's accepted proposal has a "Bookkeeping", "Bookkeeping
 * Service", or "Bookkeeping Services" heading, auto-add fixed Monthly Tasks.
 */
export async function seedMonthlyTasksFromAcceptedProposal(
  profile: ClientProfileEntry,
  proposalSlug?: string,
): Promise<ClientProfileEntry> {
  const slug = (proposalSlug || profile.activeProposalSlug).trim();
  if (!slug) {
    return profile;
  }

  try {
    const view = await fetchIgnitionProposalView(slug);
    if (!proposalHasBookkeepingServicesHeading(view.services)) {
      return profile;
    }
    return {
      ...profile,
      monthlyTasks: ensureBookkeepingMonthlyTasks(profile.monthlyTasks, slug),
    };
  } catch (error) {
    console.error("Ignition monthly task seed failed:", error);
    return profile;
  }
}
