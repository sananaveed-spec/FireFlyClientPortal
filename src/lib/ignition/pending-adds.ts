import { copyFile, mkdir, readFile, unlink, writeFile } from "fs/promises";
import path from "path";
import { isGraphAppOnlyConfigured } from "@/lib/azure-graph-app";
import {
  normalizePortalLinkType,
  type PortalLinkType,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import {
  appendDashboardNotifications,
  dismissDashboardNotification,
} from "@/lib/dashboard-notifications";
import {
  appendHistory,
  getDataDir,
  readClientProfiles,
  writeClientProfiles,
} from "@/lib/data-store";
import {
  createProfilesForAcceptedIgnitionClients,
  matchIgnitionClientToProfile,
  type IgnitionProposalRecord,
} from "@/lib/ignition/proposals";
import { seedOneTimeTasksFromAcceptedProposal } from "@/lib/ignition/seed-one-time-tasks";
import {
  buildSharePointMemberEmails,
  createClientSharePointSiteAppOnly,
} from "@/lib/sharepoint-graph";

export type IgnitionPendingAdd = {
  id: string;
  ignitionClientSlug: string;
  clientName: string;
  clientEmail: string;
  proposalSlug: string;
  proposalName: string;
  proposalReference: string;
  proposalViewUrl: string;
  proposalPdfUrl: string;
  /** When set, Add only configures portal on this existing profile. */
  existingProfileId: string;
  notificationId: string;
  createdAt: string;
  completedAt: string;
};

export type IgnitionPendingAddsStore = {
  pending: IgnitionPendingAdd[];
};

export type IgnitionPortalChoice = PortalLinkType | "skip";

function pendingAddsPath() {
  return path.join(getDataDir(), "ignition-pending-adds.json");
}

async function ensureDataDir() {
  await mkdir(getDataDir(), { recursive: true });
}

async function writeJsonFile(filePath: string, value: unknown) {
  const tempPath = `${filePath}.${process.pid}.tmp`;
  const payload = `${JSON.stringify(value, null, 2)}\n`;
  await writeFile(tempPath, payload, "utf8");
  await copyFile(tempPath, filePath);
  await unlink(tempPath).catch(() => undefined);
}

function isMissingFile(error: unknown) {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    (error as { code?: string }).code === "ENOENT"
  );
}

export function createEmptyIgnitionPendingAddsStore(): IgnitionPendingAddsStore {
  return { pending: [] };
}

export function createIgnitionPendingAddId() {
  return `pending-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function normalizePendingAdd(value: unknown): IgnitionPendingAdd | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<IgnitionPendingAdd>;
  const id = String(record.id ?? "").trim();
  const ignitionClientSlug = String(record.ignitionClientSlug ?? "").trim();
  if (!id || !ignitionClientSlug) {
    return null;
  }
  return {
    id,
    ignitionClientSlug,
    clientName: String(record.clientName ?? "").trim(),
    clientEmail: String(record.clientEmail ?? "").trim(),
    proposalSlug: String(record.proposalSlug ?? "").trim(),
    proposalName: String(record.proposalName ?? "").trim(),
    proposalReference: String(record.proposalReference ?? "").trim(),
    proposalViewUrl: String(record.proposalViewUrl ?? "").trim(),
    proposalPdfUrl: String(record.proposalPdfUrl ?? "").trim(),
    existingProfileId: String(record.existingProfileId ?? "").trim(),
    notificationId: String(record.notificationId ?? "").trim(),
    createdAt:
      String(record.createdAt ?? "").trim() || new Date().toISOString(),
    completedAt: String(record.completedAt ?? "").trim(),
  };
}

export function normalizeIgnitionPendingAddsStore(
  value: unknown,
): IgnitionPendingAddsStore {
  const record =
    typeof value === "object" && value !== null
      ? (value as { pending?: unknown })
      : {};
  const pending = Array.isArray(record.pending)
    ? record.pending
        .map(normalizePendingAdd)
        .filter((item): item is IgnitionPendingAdd => Boolean(item))
    : [];
  return { pending };
}

export async function readIgnitionPendingAddsStore(): Promise<IgnitionPendingAddsStore> {
  await ensureDataDir();
  try {
    const raw = await readFile(pendingAddsPath(), "utf8");
    return normalizeIgnitionPendingAddsStore(
      JSON.parse(raw.replace(/^\uFEFF/, "")),
    );
  } catch (error) {
    if (isMissingFile(error)) {
      return createEmptyIgnitionPendingAddsStore();
    }
    throw error;
  }
}

export async function writeIgnitionPendingAddsStore(
  store: IgnitionPendingAddsStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(pendingAddsPath(), store);
}

export function listOpenIgnitionPendingAdds(
  store: IgnitionPendingAddsStore,
): IgnitionPendingAdd[] {
  return store.pending
    .filter((item) => !item.completedAt)
    .sort(
      (a, b) =>
        new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
    );
}

function proposalFromPending(item: IgnitionPendingAdd): IgnitionProposalRecord {
  const now = item.createdAt || new Date().toISOString();
  return {
    slug: item.proposalSlug || `pending-${item.id}`,
    clientSlug: item.ignitionClientSlug,
    clientName: item.clientName,
    name: item.proposalName || item.proposalReference || "Accepted proposal",
    referenceNumber: item.proposalReference,
    state: "accepted",
    updatedAt: now,
    createdAt: now,
    startDate: "",
    minimumContractLength: null,
    endsOn: "",
    viewUrl: item.proposalViewUrl,
    pdfUrl: item.proposalPdfUrl,
  };
}

/**
 * Queue an admin “Add to system” step after proposal.accepted.
 * Skips if an open pending already exists for this Ignition client.
 */
export async function enqueueIgnitionPendingAdd(input: {
  ignitionClientSlug: string;
  clientName: string;
  clientEmail: string;
  proposal: Pick<
    IgnitionProposalRecord,
    | "slug"
    | "name"
    | "referenceNumber"
    | "viewUrl"
    | "pdfUrl"
  >;
  existingProfileId?: string;
  source?: string;
}): Promise<IgnitionPendingAdd | null> {
  const ignitionClientSlug = input.ignitionClientSlug.trim();
  if (!ignitionClientSlug) {
    return null;
  }

  const store = await readIgnitionPendingAddsStore();
  const alreadyOpen = store.pending.find(
    (item) =>
      !item.completedAt && item.ignitionClientSlug === ignitionClientSlug,
  );
  if (alreadyOpen) {
    return alreadyOpen;
  }

  const clientName = input.clientName.trim() || ignitionClientSlug;
  const proposalLabel =
    input.proposal.referenceNumber.trim() ||
    input.proposal.name.trim() ||
    input.proposal.slug.trim();
  const pendingId = createIgnitionPendingAddId();

  const createdNotes = await appendDashboardNotifications([
    {
      kind: "ignition_proposal_accepted",
      title: "Ignition proposal accepted",
      message: `${clientName}: proposal ${proposalLabel || "accepted"}. Add to system to finish Client Info${
        input.existingProfileId ? " portal setup" : ""
      }.`,
      clientName,
      clientProfileId: input.existingProfileId?.trim() || "",
      pendingAddId: pendingId,
      action: "add_to_system",
    },
  ]);
  const notificationId = createdNotes[0]?.id ?? "";

  const pending: IgnitionPendingAdd = {
    id: pendingId,
    ignitionClientSlug,
    clientName,
    clientEmail: input.clientEmail.trim(),
    proposalSlug: input.proposal.slug.trim(),
    proposalName: input.proposal.name.trim(),
    proposalReference: input.proposal.referenceNumber.trim(),
    proposalViewUrl: input.proposal.viewUrl.trim(),
    proposalPdfUrl: input.proposal.pdfUrl.trim(),
    existingProfileId: input.existingProfileId?.trim() || "",
    notificationId,
    createdAt: new Date().toISOString(),
    completedAt: "",
  };

  await writeIgnitionPendingAddsStore({
    pending: [pending, ...store.pending].slice(0, 200),
  });

  return pending;
}

async function applyPortalChoiceToProfile(
  profile: ClientProfileEntry,
  portalChoice: IgnitionPortalChoice,
  portalLinkInput: string,
): Promise<{ profile: ClientProfileEntry; notes: string }> {
  if (portalChoice === "skip") {
    return {
      profile: {
        ...profile,
        portalLink: profile.portalLink,
        portalLinkType: profile.portalLinkType || "",
      },
      notes: "Portal skipped.",
    };
  }

  if (portalChoice === "other") {
    const link = portalLinkInput.trim();
    return {
      profile: {
        ...profile,
        portalLink: link,
        portalLinkType: "other",
      },
      notes: link ? `Other portal link saved.` : "Other portal chosen; link left blank.",
    };
  }

  // firefly-sharepoint
  if (profile.portalLink.trim()) {
    return {
      profile: {
        ...profile,
        portalLinkType: "firefly-sharepoint",
      },
      notes: "FireFly SharePoint already linked.",
    };
  }

  if (!isGraphAppOnlyConfigured()) {
    throw new Error(
      "AZURE_CLIENT_SECRET is not set. Cannot create FireFly SharePoint yet — choose Other or Skip.",
    );
  }

  const memberEmails = buildSharePointMemberEmails({
    ownerEmail: profile.ownerEmail,
    pocEmail: profile.pocEmail,
  });
  const created = await createClientSharePointSiteAppOnly(
    profile.legalName,
    memberEmails,
  );
  const warn =
    created.memberWarnings.length > 0
      ? ` Member notes: ${created.memberWarnings.slice(0, 3).join("; ")}`
      : "";
  return {
    profile: {
      ...profile,
      portalLink: created.webUrl,
      portalLinkType: "firefly-sharepoint",
    },
    notes: `SharePoint created (${created.webUrl}).${warn}`,
  };
}

export type CompleteIgnitionPendingAddResult = {
  profileId: string;
  legalName: string;
  portalNotes: string;
  createdNewProfile: boolean;
};

/**
 * Admin completes “Add to system”: create Client Info if needed, then set portal.
 */
export async function completeIgnitionPendingAdd(input: {
  pendingId: string;
  portalChoice: IgnitionPortalChoice;
  portalLink?: string;
  actorEmail?: string;
  actorName?: string;
}): Promise<CompleteIgnitionPendingAddResult> {
  const pendingId = input.pendingId.trim();
  const portalChoice = input.portalChoice;
  if (
    portalChoice !== "firefly-sharepoint" &&
    portalChoice !== "other" &&
    portalChoice !== "skip"
  ) {
    throw new Error('portalChoice must be "firefly-sharepoint", "other", or "skip".');
  }

  const store = await readIgnitionPendingAddsStore();
  const index = store.pending.findIndex((item) => item.id === pendingId);
  if (index < 0) {
    throw new Error("Pending add not found.");
  }
  const pending = store.pending[index]!;
  if (pending.completedAt) {
    throw new Error("This proposal was already added to the system.");
  }

  let profiles = await readClientProfiles();
  let profile: ClientProfileEntry | undefined;
  let createdNewProfile = false;

  if (pending.existingProfileId) {
    profile = profiles.find((item) => item.id === pending.existingProfileId);
  }
  if (!profile) {
    profile = profiles.find(
      (item) =>
        matchIgnitionClientToProfile(item, [
          {
            id: pending.ignitionClientSlug,
            name: pending.clientName,
            email: pending.clientEmail,
          },
        ])?.id === pending.ignitionClientSlug,
    );
  }

  if (!profile) {
    const proposal = proposalFromPending(pending);
    const ensured = createProfilesForAcceptedIgnitionClients(
      profiles,
      [
        {
          id: pending.ignitionClientSlug,
          name: pending.clientName,
          email: pending.clientEmail,
        },
      ],
      [proposal],
    );
    profiles = ensured.profiles;
    const created = ensured.created[0];
    if (!created) {
      throw new Error("Could not create Client Info for this Ignition client.");
    }
    profile = profiles.find((item) => item.id === created.profileId);
    createdNewProfile = true;
    if (!profile) {
      throw new Error("Created Client Info could not be loaded.");
    }
  }

  const applied = await applyPortalChoiceToProfile(
    profile,
    portalChoice,
    input.portalLink ?? "",
  );
  profile = applied.profile;
  profile = await seedOneTimeTasksFromAcceptedProposal(
    profile,
    pending.proposalSlug || profile.activeProposalSlug,
  );

  const nextProfiles = profiles.map((item) =>
    item.id === profile!.id ? profile! : item,
  );
  await writeClientProfiles(nextProfiles);

  const completedAt = new Date().toISOString();
  const nextPending = [...store.pending];
  nextPending[index] = { ...pending, completedAt };
  await writeIgnitionPendingAddsStore({ pending: nextPending });

  if (pending.notificationId) {
    try {
      await dismissDashboardNotification(pending.notificationId);
    } catch {
      // Non-fatal if notification already dismissed.
    }
  }

  try {
    await appendHistory({
      employeeName: input.actorName?.trim() || "Admin",
      employeeEmail: input.actorEmail?.trim() || "",
      project: profile.legalName,
      entityType: "clientProfile",
      entityLabel: profile.legalName,
      action: createdNewProfile ? "created" : "updated",
      summary: createdNewProfile
        ? `Client added to system from Ignition proposal: ${profile.legalName}`
        : `Portal set from Ignition proposal accept: ${profile.legalName}`,
      details: [
        pending.proposalReference
          ? `Proposal: ${pending.proposalReference}`
          : "",
        pending.ignitionClientSlug
          ? `Ignition client: ${pending.ignitionClientSlug}`
          : "",
        `Portal: ${portalChoice}`,
        applied.notes,
      ]
        .filter(Boolean)
        .join("\n"),
      comment: "Completed via Ignition Add to system.",
    });
  } catch (error) {
    console.error("Ignition pending-add history failed:", error);
  }

  return {
    profileId: profile.id,
    legalName: profile.legalName,
    portalNotes: applied.notes,
    createdNewProfile,
  };
}

export function parsePortalChoice(value: unknown): IgnitionPortalChoice | null {
  const raw = String(value ?? "").trim().toLowerCase();
  if (raw === "firefly-sharepoint" || raw === "sharepoint") {
    return "firefly-sharepoint";
  }
  if (raw === "other") {
    return "other";
  }
  if (raw === "skip" || raw === "") {
    return "skip";
  }
  const normalized = normalizePortalLinkType(raw);
  return normalized || null;
}
