import {
  ensureMonthlyTasksFromServices,
  isMonthlyBillingMode,
} from "@/lib/monthly-tasks";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import {
  fetchIgnitionProposalView,
  type IgnitionProposalServiceView,
} from "@/lib/ignition/proposals";

export function monthlyServicesFromProposal(
  services: ReadonlyArray<IgnitionProposalServiceView>,
): Array<{ name: string }> {
  return services
    .filter((service) => isMonthlyBillingMode(service.billingMode || ""))
    .map((service) => ({ name: service.name }));
}

/**
 * When an accepted Ignition proposal includes monthly-billed services,
 * auto-add Monthly Tasks for each (empty list when no proposal / no matches).
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
    const monthlyServices = monthlyServicesFromProposal(view.services);
    if (monthlyServices.length === 0) {
      return profile;
    }
    return {
      ...profile,
      monthlyTasks: ensureMonthlyTasksFromServices(
        profile.monthlyTasks,
        monthlyServices,
        slug,
      ),
    };
  } catch (error) {
    console.error("Ignition monthly task seed failed:", error);
    return profile;
  }
}
