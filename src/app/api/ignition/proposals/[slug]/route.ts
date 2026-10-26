import { NextResponse } from "next/server";
import { fetchIgnitionProposalView } from "@/lib/ignition/proposals";
import { requirePortalUser } from "@/lib/requireEditor";

type RouteContext = {
  params: Promise<{ slug: string }>;
};

export async function GET(request: Request, context: RouteContext) {
  const auth = await requirePortalUser(request);
  if (auth instanceof NextResponse) {
    return auth;
  }

  try {
    const { slug } = await context.params;
    const proposal = await fetchIgnitionProposalView(slug);
    return NextResponse.json({ proposal });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load Ignition proposal.";
    const status = /not connected|reconnect/i.test(message) ? 400 : 500;
    return NextResponse.json({ error: message }, { status });
  }
}
