import {
  ensureBookkeepingMonthlyTasks,
  ensureFinancialReportingMonthlyTasks,
} from "@/lib/monthly-tasks";
import type { ClientProfileEntry } from "@/lib/client-profiles";
import { fetchIgnitionProposalView } from "@/lib/ignition/proposals";
import {
  proposalHasBookkeepingServicesHeading,
  proposalHasFinancialReportingServicesHeading,
} from "@/lib/one-time-tasks";

/**
 * Called when Resource Allocation is saved (not on proposal accept).
 * Seeds fixed Monthly Tasks from proposal description/group headings:
 * Bookkeeping / Bookkeeping Service(s), and Financial Reporting Service(s).
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
    let monthlyTasks = profile.monthlyTasks;
    if (proposalHasBookkeepingServicesHeading(view.services)) {
      monthlyTasks = ensureBookkeepingMonthlyTasks(monthlyTasks, slug);
    }
    if (proposalHasFinancialReportingServicesHeading(view.services)) {
      monthlyTasks = ensureFinancialReportingMonthlyTasks(monthlyTasks, slug);
    }
    if (monthlyTasks === profile.monthlyTasks) {
      return profile;
    }
    return {
      ...profile,
      monthlyTasks,
    };
  } catch (error) {
    console.error("Ignition monthly task seed failed:", error);
    return profile;
  }
}
