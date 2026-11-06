import { NextResponse } from "next/server";
import {
  completeIgnitionPendingAdd,
  listOpenIgnitionPendingAdds,
  parsePortalChoice,
  readIgnitionPendingAddsStore,
} from "@/lib/ignition/pending-adds";
import { requireAdmin, requirePortalUser } from "@/lib/requireEditor";
import { readPortalUsers } from "@/lib/portal-users-store";
import { normalizeEmail } from "@/lib/portal-users";

export async function GET(request: Request) {
  const gate = await requirePortalUser(request);
  if (gate instanceof NextResponse) {
    return gate;
  }

  try {
    const store = await readIgnitionPendingAddsStore();
    return NextResponse.json({
      pending: listOpenIgnitionPendingAdds(store),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load Ignition pending adds.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json().catch(() => ({}))) as {
      action?: string;
      pendingId?: string;
      portalChoice?: string;
      portalLink?: string;
    };
    const action = String(body.action ?? "").trim().toLowerCase();
    if (action !== "add") {
      return NextResponse.json(
        { error: 'Expected action "add".' },
        { status: 400 },
      );
    }

    const pendingId = String(body.pendingId ?? "").trim();
    const portalChoice = parsePortalChoice(body.portalChoice);
    if (!pendingId || !portalChoice) {
      return NextResponse.json(
        {
          error:
            'pendingId and portalChoice ("firefly-sharepoint", "other", or "skip") are required.',
        },
        { status: 400 },
      );
    }

    const actorEmail = normalizeEmail(
      request.headers.get("x-user-email") ?? "",
    );
    const users = await readPortalUsers();
    const actor = users.find(
      (user) => normalizeEmail(user.email) === actorEmail,
    );

    const result = await completeIgnitionPendingAdd({
      pendingId,
      portalChoice,
      portalLink: String(body.portalLink ?? ""),
      actorEmail,
      actorName: actor?.name || actorEmail,
    });

    return NextResponse.json({
      ok: true,
      ...result,
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to add Ignition client to the system.";
    return NextResponse.json({ error: message }, { status: 400 });
  }
}
