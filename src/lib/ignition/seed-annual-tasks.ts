import {
  ensureAnnualTasksFromTaxServices,
  isTaxServiceGroupName,
} from "@/lib/annual-tasks";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import {
  fetchIgnitionProposalView,
  type IgnitionProposalServiceView,
} from "@/lib/ignition/proposals";

export function taxServicesFromProposal(
  services: ReadonlyArray<IgnitionProposalServiceView>,
): Array<{ name: string; descriptionHtml: string }> {
  const fromGroups = services.filter((service) =>
    isTaxServiceGroupName(service.groupName || ""),
  );
  if (fromGroups.length > 0) {
    return fromGroups.map((service) => ({
      name: service.name,
      descriptionHtml: service.descriptionHtml,
    }));
  }
  // Fallback: service name itself looks like tax compliance work.
  return services
    .filter((service) =>
      /tax\s*(return|compliance|services?)|s-?corporation|c-?corporation/i.test(
        service.name,
      ),
    )
    .map((service) => ({
      name: service.name,
      descriptionHtml: service.descriptionHtml,
    }));
}

/**
 * When an accepted Ignition proposal includes Tax / Tax Compliance services,
 * auto-add Annual Tasks for each service (with due dates from descriptions).
 */
export async function seedAnnualTasksFromAcceptedProposal(
  profile: ClientProfileEntry,
  proposalSlug?: string,
): Promise<ClientProfileEntry> {
  const slug = (proposalSlug || profile.activeProposalSlug).trim();
  if (!slug) {
    return profile;
  }

  try {
    const view = await fetchIgnitionProposalView(slug);
    const taxServices = taxServicesFromProposal(view.services);
    if (taxServices.length === 0) {
      return profile;
    }
    return {
      ...profile,
      annualTasks: ensureAnnualTasksFromTaxServices(
        profile.annualTasks,
        taxServices,
        slug,
      ),
    };
  } catch (error) {
    console.error("Ignition annual task seed failed (tax services):", error);
    return profile;
  }
}
