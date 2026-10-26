import { NextResponse } from "next/server";
import {
  readClientProfiles,
  readEmailTemplates,
  readHolidays,
  readIgnitionEndingSoonEmailsFile,
  writeIgnitionEndingSoonEmailsFile,
} from "@/lib/data-store";
import { normalizeEmailTemplates } from "@/lib/email-templates";
import { readIgnitionClientsStore } from "@/lib/ignition/clients";
import { readIgnitionEndingSoonStore } from "@/lib/ignition/proposals";
import {
  applyEndingSoonEmailPlansToStore,
  buildEndingSoonEmailPlans,
  buildEndingSoonEmailStatusBySlug,
  collectEndingSoonTeamEmails,
  ENDING_SOON_CLIENT_EMAIL_TEMPLATE_ID,
  ENDING_SOON_EMAILS_DRY_RUN,
  ENDING_SOON_TEAM_EMAIL_TEMPLATE_ID,
  endingSoonEmailSendersLabel,
  isEndingSoonEmailSender,
  normalizeEndingSoonEmailsStore,
  type EndingSoonProposalInput,
} from "@/lib/ignition-ending-soon-emails";
import { normalizeEmail } from "@/lib/portal-users";
import { readPortalUsers } from "@/lib/portal-users-store";
import { requireAdmin, requirePortalUser } from "@/lib/requireEditor";

function requireEndingSoonSender(
  request: Request,
): NextResponse | { email: string } {
  const email = normalizeEmail(request.headers.get("x-user-email") ?? "");
  if (!isEndingSoonEmailSender(email)) {
    return NextResponse.json(
      {
        error: `Only ${endingSoonEmailSendersLabel()} can update Ending Soon email status.`,
      },
      { status: 403 },
    );
  }
  return { email };
}

async function loadProposalInputs(): Promise<{
  proposals: EndingSoonProposalInput[];
  ignitionEmailsByClientSlug: Record<string, string>;
}> {
  const [endingSoonStore, clientsStore] = await Promise.all([
    readIgnitionEndingSoonStore(),
    readIgnitionClientsStore(),
  ]);
  const ignitionEmailsByClientSlug: Record<string, string> = {};
  for (const client of clientsStore.clients) {
    const email = normalizeEmail(client.email);
    if (email) {
      ignitionEmailsByClientSlug[client.id] = email;
    }
  }
  const proposals: EndingSoonProposalInput[] = endingSoonStore.proposals.map(
    (proposal) => ({
      slug: proposal.slug,
      clientSlug: proposal.clientSlug,
      clientName: proposal.clientName,
      name: proposal.name,
      referenceNumber: proposal.referenceNumber,
      endsOn: proposal.endsOn,
    }),
  );
  return { proposals, ignitionEmailsByClientSlug };
}

export async function GET(request: Request) {
  const gate = await requirePortalUser(request);
  if (gate instanceof NextResponse) {
    return gate;
  }

  try {
    const [emailsRaw, { proposals }] = await Promise.all([
      readIgnitionEndingSoonEmailsFile(),
      loadProposalInputs(),
    ]);
    const store = normalizeEndingSoonEmailsStore(emailsRaw);

    return NextResponse.json({
      dryRun: true,
      sendsDisabled: true,
      entries: store.entries,
      teamStatusBySlug: buildEndingSoonEmailStatusBySlug(
        proposals,
        store,
        "team",
      ),
      clientStatusBySlug: buildEndingSoonEmailStatusBySlug(
        proposals,
        store,
        "client",
      ),
      allowedSenders: endingSoonEmailSendersLabel(),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load Ending Soon email status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

/**
 * Status/logic only: marks Team + Client as Scheduled. Never sends mail.
 */
export async function POST(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  const senderGate = requireEndingSoonSender(request);
  if (senderGate instanceof NextResponse) {
    return senderGate;
  }

  try {
    const body = (await request.json().catch(() => ({}))) as {
      action?: string;
    };
    const action = String(body.action ?? "ensure").trim().toLowerCase();

    if (action !== "ensure") {
      return NextResponse.json(
        {
          error:
            'Only action "ensure" is supported (status/logic only — no send).',
        },
        { status: 400 },
      );
    }

    if (!ENDING_SOON_EMAILS_DRY_RUN) {
      return NextResponse.json(
        {
          error:
            "Ending Soon dry-run is off unexpectedly. Refusing to send mail.",
        },
        { status: 500 },
      );
    }

    const [
      emailsRaw,
      templatesRaw,
      profilesRaw,
      portalUsers,
      holidays,
      { proposals, ignitionEmailsByClientSlug },
    ] = await Promise.all([
      readIgnitionEndingSoonEmailsFile(),
      readEmailTemplates(),
      readClientProfiles(),
      readPortalUsers(),
      readHolidays(),
      loadProposalInputs(),
    ]);

    const templates = normalizeEmailTemplates(templatesRaw);
    const store = normalizeEndingSoonEmailsStore(emailsRaw);
    const teamTemplate = templates.templates.find(
      (item) => item.id === ENDING_SOON_TEAM_EMAIL_TEMPLATE_ID,
    );
    const clientTemplate = templates.templates.find(
      (item) => item.id === ENDING_SOON_CLIENT_EMAIL_TEMPLATE_ID,
    );
    if (!teamTemplate || !clientTemplate) {
      return NextResponse.json(
        {
          error:
            "Ending Soon templates (template-19 / template-20) are missing.",
        },
        { status: 500 },
      );
    }

    const teamEmails = collectEndingSoonTeamEmails(portalUsers);
    const plans = buildEndingSoonEmailPlans({
      proposals,
      teamEmails,
      profiles: profilesRaw,
      store,
      ignitionEmailsByClientSlug,
      teamTemplate,
      clientTemplate,
      holidays,
    });

    const nextStore = applyEndingSoonEmailPlansToStore(store, plans);
    await writeIgnitionEndingSoonEmailsFile(nextStore);

    return NextResponse.json({
      dryRun: true,
      sendsDisabled: true,
      message:
        "Team and Client emails marked Scheduled for 9:00 PM Pakistan time. Nothing was sent.",
      plannedCount: plans.length,
      teamPlanCount: plans.filter((plan) => plan.audience === "team").length,
      clientPlanCount: plans.filter((plan) => plan.audience === "client")
        .length,
      teamStatusBySlug: buildEndingSoonEmailStatusBySlug(
        proposals,
        nextStore,
        "team",
      ),
      clientStatusBySlug: buildEndingSoonEmailStatusBySlug(
        proposals,
        nextStore,
        "client",
      ),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to update Ending Soon email status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
