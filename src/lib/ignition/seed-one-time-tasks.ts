import type { ClientProfileEntry } from "@/lib/client-profiles";
import {
  ensureQboSetupTaskFromBookkeeping,
  hasQboSetupOneTimeTask,
  proposalHasBookkeepingServicesHeading,
} from "@/lib/one-time-tasks";
import { fetchIgnitionProposalView } from "@/lib/ignition/proposals";

/**
 * When an accepted Ignition proposal has a "Bookkeeping Services" (or
 * "Bookkeeping Service") heading, auto-add the QBO Setup one-time task.
 */
export async function seedOneTimeTasksFromAcceptedProposal(
  profile: ClientProfileEntry,
  proposalSlug?: string,
): Promise<ClientProfileEntry> {
  const slug = (proposalSlug || profile.activeProposalSlug).trim();
  if (!slug) {
    return profile;
  }
  if (hasQboSetupOneTimeTask(profile.oneTimeTasks)) {
    return profile;
  }

  try {
    const view = await fetchIgnitionProposalView(slug);
    const hasHeading = proposalHasBookkeepingServicesHeading(view.services);
    return {
      ...profile,
      oneTimeTasks: ensureQboSetupTaskFromBookkeeping(
        profile.oneTimeTasks,
        hasHeading,
        slug,
      ),
    };
  } catch (error) {
    console.error(
      "Ignition one-time task seed failed (Bookkeeping Services / QBO Setup):",
      error,
    );
    return profile;
  }
}
