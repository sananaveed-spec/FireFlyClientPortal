import { ensureIncomeTaxComplianceAnnualTasks } from "@/lib/annual-tasks";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import { fetchIgnitionProposalView } from "@/lib/ignition/proposals";
import { proposalHasIncomeTaxComplianceServicesHeading } from "@/lib/one-time-tasks";

/**
 * When an accepted Ignition proposal has an "Income Tax Compliance Service"
 * or "Income Tax Compliance Services" heading (often in description HTML),
 * auto-add the fixed Annual Tasks. Due dates are left blank for users to set.
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
    if (!proposalHasIncomeTaxComplianceServicesHeading(view.services)) {
      return profile;
    }
    return {
      ...profile,
      annualTasks: ensureIncomeTaxComplianceAnnualTasks(
        profile.annualTasks,
        slug,
      ),
    };
  } catch (error) {
    console.error(
      "Ignition annual task seed failed (Income Tax Compliance):",
      error,
    );
    return profile;
  }
}
