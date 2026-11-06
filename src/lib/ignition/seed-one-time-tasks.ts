import type { ClientProfileEntry } from "@/lib/client-profiles";
import {
  ensureQboSetupTaskFromBookkeeping,
  hasQboSetupOneTimeTask,
  proposalServicesIncludeBookkeeping,
} from "@/lib/one-time-tasks";
import { fetchIgnitionProposalView } from "@/lib/ignition/proposals";

/**
 * When an accepted Ignition proposal includes bookkeeping services,
 * auto-add the single QBO Setup one-time task (if not already present).
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
    const hasBookkeeping = proposalServicesIncludeBookkeeping(view.services);
    return {
      ...profile,
      oneTimeTasks: ensureQboSetupTaskFromBookkeeping(
        profile.oneTimeTasks,
        hasBookkeeping,
        slug,
      ),
    };
  } catch (error) {
    console.error(
      "Ignition one-time task seed failed (bookkeeping / QBO Setup):",
      error,
    );
    return profile;
  }
}
