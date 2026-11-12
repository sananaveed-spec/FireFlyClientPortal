import { ensureGstHstTaxComplianceQuarterlyTasks } from "@/lib/quarterly-tasks";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import { fetchIgnitionProposalView } from "@/lib/ignition/proposals";
import { proposalHasGstHstTaxComplianceServicesHeading } from "@/lib/one-time-tasks";

/**
 * Called when Resource Allocation is saved (not on proposal accept).
 * If the accepted proposal has GST/HST Tax Compliance Service(s) heading,
 * auto-add the fixed Quarterly Task (due date left blank for users).
 */
export async function seedQuarterlyTasksFromAcceptedProposal(
  profile: ClientProfileEntry,
  proposalSlug?: string,
): Promise<ClientProfileEntry> {
  const slug = (proposalSlug || profile.activeProposalSlug).trim();
  if (!slug) {
    return profile;
  }

  try {
    const view = await fetchIgnitionProposalView(slug);
    if (!proposalHasGstHstTaxComplianceServicesHeading(view.services)) {
      return profile;
    }
    return {
      ...profile,
      quarterlyTasks: ensureGstHstTaxComplianceQuarterlyTasks(
        profile.quarterlyTasks,
        slug,
      ),
    };
  } catch (error) {
    console.error(
      "Ignition quarterly task seed failed (GST/HST Tax Compliance):",
      error,
    );
    return profile;
  }
}
