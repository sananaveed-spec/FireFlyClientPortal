import { NextResponse } from "next/server";
import {
  readIgnitionClientsStore,
  syncIgnitionClients,
} from "@/lib/ignition/clients";
import {
  buildIgnitionClientInvoiceSummaries,
  readIgnitionInvoicesStore,
} from "@/lib/ignition/invoices";
import { readIgnitionEndingSoonStore } from "@/lib/ignition/proposals";
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

async function buildResponsePayload(
  store: Awaited<ReturnType<typeof readIgnitionClientsStore>>,
) {
  const endingSoonStore = await readIgnitionEndingSoonStore();
  const invoicesStore = await readIgnitionInvoicesStore();
  const clientNameBySlug = clientNameBySlugFromStore(store);
  return {
    syncedAt: store.syncedAt || null,
    endingSoonSyncedAt: endingSoonStore.syncedAt || null,
    invoicesSyncedAt: invoicesStore.syncedAt || null,
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
    newClientsFromAcceptedProposals: [] as Array<{
      profileId: string;
      uniqueId: string;
      legalName: string;
      ignitionClientSlug: string;
      proposalSlug: string;
      proposalReference: string;
    }>,
  };
}

export async function GET(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const store = await readIgnitionClientsStore();
    return NextResponse.json(await buildResponsePayload(store));
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load synced Ignition clients.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const store = await syncIgnitionClients();
    const payload = await buildResponsePayload(store);
    return NextResponse.json({
      ...payload,
      newClientsFromAcceptedProposals:
        store.newClientsFromAcceptedProposals ?? [],
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to sync Ignition clients.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
