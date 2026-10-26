import { NextResponse } from "next/server";
import { readIgnitionClientsStore } from "@/lib/ignition/clients";
import {
  buildIgnitionClientInvoiceSummaries,
  readIgnitionInvoicesStore,
  syncIgnitionInvoices,
} from "@/lib/ignition/invoices";
import {
  IGNITION_ENDING_SOON_MONTHS,
  readIgnitionEndingSoonStore,
  syncIgnitionEndingSoon,
} from "@/lib/ignition/proposals";
import {
  ignitionConnectionStatus,
  readIgnitionTokens,
} from "@/lib/ignition/tokens";
import { requireAdmin } from "@/lib/requireEditor";

function mapEndingSoonProposals(
  store: Awaited<ReturnType<typeof readIgnitionEndingSoonStore>>,
) {
  return store.proposals.map((proposal) => ({
    slug: proposal.slug,
    clientSlug: proposal.clientSlug,
    clientName: proposal.clientName,
    name: proposal.name,
    referenceNumber: proposal.referenceNumber,
    state: proposal.state,
    startDate: proposal.startDate,
    minimumContractLength: proposal.minimumContractLength,
    endsOn: proposal.endsOn,
    viewUrl: proposal.viewUrl,
  }));
}

function clientNameBySlugFromStore(
  store: Awaited<ReturnType<typeof readIgnitionClientsStore>>,
): Record<string, string> {
  const map: Record<string, string> = {};
  for (const client of store.clients) {
    map[client.id] = client.name;
  }
  return map;
}

export async function GET(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const tokens = await readIgnitionTokens();
    const status = ignitionConnectionStatus(tokens);
    const store = await readIgnitionClientsStore();
    let endingSoonStore = await readIgnitionEndingSoonStore();
    let invoicesStore = await readIgnitionInvoicesStore();
    const clientNameBySlug = clientNameBySlugFromStore(store);

    // Hydrate / rebuild when empty or when the month window setting changed.
    const windowMismatch =
      endingSoonStore.windowMonths !== IGNITION_ENDING_SOON_MONTHS;
    if (
      status.connected &&
      (endingSoonStore.proposals.length === 0 || windowMismatch)
    ) {
      try {
        endingSoonStore = await syncIgnitionEndingSoon();
      } catch (error) {
        console.error("Ignition ending-soon hydrate failed:", error);
      }
    }

    if (status.connected && invoicesStore.invoices.length === 0) {
      try {
        invoicesStore = await syncIgnitionInvoices(clientNameBySlug);
      } catch (error) {
        console.error("Ignition invoices hydrate failed:", error);
      }
    }

    return NextResponse.json({
      connected: status.connected,
      configured: status.configured,
      expiresAt: status.expiresAt,
      syncedAt: store.syncedAt || null,
      endingSoonSyncedAt: endingSoonStore.syncedAt || null,
      invoicesSyncedAt: invoicesStore.syncedAt || null,
      clientCount: store.clients.length,
      clients: store.clients.map((client) => ({
        id: client.id,
        name: client.name,
        email: client.email,
        status: client.status,
      })),
      endingSoon: mapEndingSoonProposals(endingSoonStore),
      invoiceSummaries: buildIgnitionClientInvoiceSummaries(
        invoicesStore.invoices,
        clientNameBySlug,
      ),
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load Ignition status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
