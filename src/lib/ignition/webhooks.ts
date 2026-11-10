import { createHmac, timingSafeEqual } from "crypto";
import { copyFile, mkdir, readFile, unlink, writeFile } from "fs/promises";
import path from "path";
import {
  appendDashboardNotifications,
} from "@/lib/dashboard-notifications";
import {
  getDataDir,
  readClientProfiles,
  writeClientProfiles,
} from "@/lib/data-store";
import {
  getIgnitionConfig,
  getIgnitionPublicBaseUrl,
} from "@/lib/ignition/config";
export { getIgnitionPublicBaseUrl } from "@/lib/ignition/config";
import { enqueueIgnitionPendingAdd } from "@/lib/ignition/pending-adds";
import {
  applyActiveProposalToProfile,
  matchIgnitionClientToProfile,
  normalizeIgnitionProposal,
  pickActiveProposalByClient,
  type IgnitionProposalRecord,
} from "@/lib/ignition/proposals";
import { seedClientEngagementTasksFromProposal } from "@/lib/ignition/seed-client-engagement-tasks";
import { getValidAccessToken } from "@/lib/ignition/tokens";
import { normalizeEmail, isValidEmail } from "@/lib/portal-users";

export const IGNITION_WEBHOOK_EVENTS = ["proposal.accepted"] as const;

export type IgnitionWebhookSubscriptionRecord = {
  slug: string;
  url: string;
  events: string[];
  status: string;
  /** Signing secret — stored after create; Ignition only shows it once. */
  secret: string;
  createdAt: string;
  updatedAt: string;
};

export type IgnitionWebhooksStore = {
  subscription: IgnitionWebhookSubscriptionRecord | null;
  /** Recently processed delivery keys for idempotency. */
  processedKeys: string[];
};

function webhooksPath() {
  return path.join(getDataDir(), "ignition-webhooks.json");
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

export function getIgnitionWebhookReceiveUrl(): string {
  return `${getIgnitionPublicBaseUrl()}/api/ignition/webhooks`;
}

export function createEmptyIgnitionWebhooksStore(): IgnitionWebhooksStore {
  return { subscription: null, processedKeys: [] };
}

export async function readIgnitionWebhooksStore(): Promise<IgnitionWebhooksStore> {
  await ensureDataDir();
  try {
    const raw = await readFile(webhooksPath(), "utf8");
    const parsed = JSON.parse(raw.replace(/^\uFEFF/, "")) as Partial<IgnitionWebhooksStore>;
    const subscription =
      parsed.subscription && typeof parsed.subscription === "object"
        ? {
            slug: String(parsed.subscription.slug ?? "").trim(),
            url: String(parsed.subscription.url ?? "").trim(),
            events: Array.isArray(parsed.subscription.events)
              ? parsed.subscription.events.map((item) => String(item))
              : [],
            status: String(parsed.subscription.status ?? "").trim(),
            secret: String(parsed.subscription.secret ?? "").trim(),
            createdAt: String(parsed.subscription.createdAt ?? "").trim(),
            updatedAt: String(parsed.subscription.updatedAt ?? "").trim(),
          }
        : null;
    return {
      subscription:
        subscription && subscription.slug ? subscription : null,
      processedKeys: Array.isArray(parsed.processedKeys)
        ? parsed.processedKeys.map((item) => String(item)).filter(Boolean)
        : [],
    };
  } catch (error) {
    if (isMissingFile(error)) {
      return createEmptyIgnitionWebhooksStore();
    }
    throw error;
  }
}

export async function writeIgnitionWebhooksStore(
  store: IgnitionWebhooksStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(webhooksPath(), store);
}

function resolveSigningSecret(storeSecret: string): string {
  return (
    process.env.IGNITION_WEBHOOK_SECRET?.trim() ||
    storeSecret.trim() ||
    ""
  );
}

/**
 * Verify X-Ignition-Signature: `t=…,v1=…` over `${t}.${rawBody}`.
 * Tries several secret encodings (full whsec_, stripped, hex-decoded).
 */
export function verifyIgnitionWebhookSignature(input: {
  rawBody: string;
  signatureHeader: string;
  secret: string;
  nowMs?: number;
  toleranceSec?: number;
}): boolean {
  const secret = input.secret.trim();
  if (!secret) {
    return false;
  }
  const header = input.signatureHeader.trim();
  const parts = Object.fromEntries(
    header
      .split(",")
      .map((part) => part.trim())
      .filter(Boolean)
      .map((part) => {
        const idx = part.indexOf("=");
        if (idx <= 0) {
          return ["", ""];
        }
        return [part.slice(0, idx), part.slice(idx + 1)];
      }),
  ) as Record<string, string>;
  const timestamp = parts.t;
  const v1 = parts.v1;
  if (!timestamp || !v1) {
    return false;
  }
  const ts = Number(timestamp);
  if (!Number.isFinite(ts)) {
    return false;
  }
  const nowSec = Math.floor((input.nowMs ?? Date.now()) / 1000);
  const tolerance = input.toleranceSec ?? 300;
  if (Math.abs(nowSec - ts) > tolerance) {
    return false;
  }

  const payload = `${timestamp}.${input.rawBody}`;
  const secretCandidates: Buffer[] = [];
  secretCandidates.push(Buffer.from(secret, "utf8"));
  const stripped = secret.replace(/^whsec_/i, "");
  if (stripped && stripped !== secret) {
    secretCandidates.push(Buffer.from(stripped, "utf8"));
    if (/^[0-9a-f]+$/i.test(stripped) && stripped.length % 2 === 0) {
      secretCandidates.push(Buffer.from(stripped, "hex"));
    }
    try {
      secretCandidates.push(Buffer.from(stripped, "base64"));
    } catch {
      // ignore
    }
  }

  const expectedHexes = secretCandidates.map((key) =>
    createHmac("sha256", key).update(payload, "utf8").digest("hex"),
  );
  const expectedB64 = secretCandidates.map((key) =>
    createHmac("sha256", key).update(payload, "utf8").digest("base64"),
  );

  const provided = v1.trim();
  for (const expected of [...expectedHexes, ...expectedB64]) {
    try {
      const a = Buffer.from(provided);
      const b = Buffer.from(expected);
      if (a.length === b.length && timingSafeEqual(a, b)) {
        return true;
      }
    } catch {
      // continue
    }
  }
  return false;
}

async function ignitionApiFetch(
  pathSuffix: string,
  init: RequestInit = {},
): Promise<Response> {
  const { accessToken } = await getValidAccessToken();
  const { apiBase } = getIgnitionConfig();
  return fetch(`${apiBase}${pathSuffix}`, {
    ...init,
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${accessToken}`,
      ...(init.body ? { "Content-Type": "application/json" } : {}),
      ...(init.headers || {}),
    },
    cache: "no-store",
  });
}

export async function listIgnitionWebhookSubscriptions(): Promise<
  Array<Record<string, unknown>>
> {
  const response = await ignitionApiFetch("/webhook_subscriptions?limit=100");
  const payload = (await response.json().catch(() => ({}))) as {
    data?: unknown;
    error?: { message?: string } | string;
    message?: string;
  };
  if (!response.ok) {
    const message =
      (typeof payload.error === "object" && payload.error?.message) ||
      (typeof payload.error === "string" && payload.error) ||
      payload.message ||
      `Failed to list webhook subscriptions (${response.status}).`;
    throw new Error(message);
  }
  return Array.isArray(payload.data) ? (payload.data as Array<Record<string, unknown>>) : [];
}

export async function deleteIgnitionWebhookSubscription(
  slug: string,
): Promise<void> {
  const response = await ignitionApiFetch(
    `/webhook_subscriptions/${encodeURIComponent(slug)}`,
    { method: "DELETE" },
  );
  if (!response.ok && response.status !== 404) {
    const payload = (await response.json().catch(() => ({}))) as {
      error?: { message?: string } | string;
      message?: string;
    };
    const message =
      (typeof payload.error === "object" && payload.error?.message) ||
      (typeof payload.error === "string" && payload.error) ||
      payload.message ||
      `Failed to delete webhook (${response.status}).`;
    throw new Error(message);
  }
}

/**
 * Register (or refresh) the portal webhook subscription for proposal.accepted.
 * Stores the one-time signing secret locally.
 */
export async function ensureIgnitionWebhookSubscription(): Promise<{
  subscription: IgnitionWebhookSubscriptionRecord;
  created: boolean;
  message: string;
}> {
  const receiveUrl = getIgnitionWebhookReceiveUrl();
  const store = await readIgnitionWebhooksStore();
  const existing = await listIgnitionWebhookSubscriptions();
  const ours = existing.filter((item) => String(item.url ?? "").trim() === receiveUrl);

  // Already registered remotely and we still have the secret.
  if (ours.length > 0 && store.subscription?.secret) {
    const remote = ours[0];
    const subscription: IgnitionWebhookSubscriptionRecord = {
      slug: String(remote.slug ?? store.subscription.slug),
      url: receiveUrl,
      events: Array.isArray(remote.events)
        ? remote.events.map((item) => String(item))
        : [...IGNITION_WEBHOOK_EVENTS],
      status: String(remote.status ?? "enabled"),
      secret: store.subscription.secret,
      createdAt: store.subscription.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    await writeIgnitionWebhooksStore({
      ...store,
      subscription,
    });
    return {
      subscription,
      created: false,
      message: "Ignition webhook already registered for this server URL.",
    };
  }

  // Remote exists but secret lost — recreate so we get a new secret.
  for (const item of ours) {
    const slug = String(item.slug ?? "").trim();
    if (slug) {
      await deleteIgnitionWebhookSubscription(slug);
    }
  }

  const response = await ignitionApiFetch("/webhook_subscriptions", {
    method: "POST",
    headers: {
      "Idempotency-Key": `portal-webhook-${Date.now()}`,
    },
    body: JSON.stringify({
      url: receiveUrl,
      events: [...IGNITION_WEBHOOK_EVENTS],
    }),
  });
  const payload = (await response.json().catch(() => ({}))) as Record<
    string,
    unknown
  >;
  if (!response.ok) {
    const errorObj =
      payload.error && typeof payload.error === "object"
        ? (payload.error as { message?: string })
        : null;
    throw new Error(
      errorObj?.message ||
        (typeof payload.error === "string" ? payload.error : "") ||
        (typeof payload.message === "string" ? payload.message : "") ||
        `Failed to register Ignition webhook (${response.status}).`,
    );
  }

  const secret = String(payload.secret ?? "").trim();
  if (!secret) {
    throw new Error(
      "Ignition webhook was created but no signing secret was returned. Delete it in Ignition and try again.",
    );
  }

  const subscription: IgnitionWebhookSubscriptionRecord = {
    slug: String(payload.slug ?? "").trim(),
    url: receiveUrl,
    events: Array.isArray(payload.events)
      ? payload.events.map((item) => String(item))
      : [...IGNITION_WEBHOOK_EVENTS],
    status: String(payload.status ?? "enabled"),
    secret,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  await writeIgnitionWebhooksStore({
    ...store,
    subscription,
  });

  return {
    subscription,
    created: true,
    message:
      "Ignition webhook registered. New accepted proposals will auto-add to Client Info and notify the team.",
  };
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (typeof value === "object" && value !== null && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

function pickString(record: Record<string, unknown>, keys: string[]): string {
  for (const key of keys) {
    const value = record[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return "";
}

async function fetchIgnitionClientMatch(clientSlug: string): Promise<{
  id: string;
  name: string;
  email: string;
} | null> {
  const slug = clientSlug.trim();
  if (!slug) {
    return null;
  }
  try {
    const response = await ignitionApiFetch(
      `/clients/${encodeURIComponent(slug)}`,
    );
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      return null;
    }
    const record = asRecord(payload);
    if (!record) {
      return null;
    }
    const id = pickString(record, ["slug", "id"]) || slug;
    const name = pickString(record, ["name", "client_name"]) || id;
    let email = pickString(record, ["email", "primary_email"]);
    if (!email) {
      const contacts = record.contacts;
      if (Array.isArray(contacts)) {
        for (const contact of contacts) {
          const row = asRecord(contact);
          if (!row) continue;
          const candidate = normalizeEmail(
            pickString(row, ["email", "email_address"]),
          );
          if (candidate && isValidEmail(candidate)) {
            email = candidate;
            break;
          }
        }
      }
    }
    return { id, name, email: normalizeEmail(email) };
  } catch {
    return null;
  }
}

export type IgnitionWebhookProcessResult = {
  handled: boolean;
  skipped?: string;
  createdClients: Array<{
    profileId: string;
    legalName: string;
    uniqueId: string;
  }>;
  updatedExisting: boolean;
};

/**
 * Process a verified Ignition webhook event (proposal.accepted).
 * Does not auto-create SharePoint or Client Info — admins get an
 * “Add to system” notification and choose FireFly SharePoint / Other / Skip.
 */
export async function processIgnitionWebhookEvent(event: {
  type: string;
  occurredAt: string;
  data: unknown;
}): Promise<IgnitionWebhookProcessResult> {
  const type = event.type.trim().toLowerCase();
  if (type !== "proposal.accepted") {
    return {
      handled: false,
      skipped: `Unhandled event type: ${event.type}`,
      createdClients: [],
      updatedExisting: false,
    };
  }

  const proposal = normalizeIgnitionProposal({
    ...(asRecord(event.data) ?? {}),
    state: "accepted",
  });
  if (!proposal) {
    return {
      handled: false,
      skipped: "Webhook payload was not a usable proposal.",
      createdClients: [],
      updatedExisting: false,
    };
  }

  const store = await readIgnitionWebhooksStore();
  const dedupeKey = `${type}:${proposal.slug}:${event.occurredAt}`;
  if (store.processedKeys.includes(dedupeKey)) {
    return {
      handled: true,
      skipped: "Duplicate delivery ignored.",
      createdClients: [],
      updatedExisting: false,
    };
  }

  const clientMatch =
    (await fetchIgnitionClientMatch(proposal.clientSlug)) || {
      id: proposal.clientSlug,
      name: proposal.clientName || proposal.clientSlug,
      email: "",
    };

  const profiles = await readClientProfiles();
  const matchedProfile = profiles.find(
    (profile) =>
      matchIgnitionClientToProfile(profile, [clientMatch])?.id ===
      clientMatch.id,
  );

  let updatedExisting = false;
  if (matchedProfile) {
    const activeByClient = pickActiveProposalByClient([proposal]);
    let updatedProfile = applyActiveProposalToProfile(
      matchedProfile,
      [clientMatch],
      activeByClient,
    );
    updatedProfile = await seedClientEngagementTasksFromProposal(
      updatedProfile,
      proposal.slug,
    );
    const nextProfiles = profiles.map((profile) =>
      profile.id === updatedProfile.id ? updatedProfile : profile,
    );
    await writeClientProfiles(nextProfiles);
    updatedExisting = true;

    if (updatedProfile.portalLink.trim()) {
      await appendDashboardNotifications([
        {
          kind: "ignition_proposal_accepted",
          title: "Ignition proposal accepted",
          message: `${updatedProfile.legalName}: proposal ${
            proposal.referenceNumber || proposal.name || proposal.slug
          } was accepted. Client Info is up to date.`,
          clientName: updatedProfile.legalName,
          clientProfileId: updatedProfile.id,
        },
      ]);
    } else {
      await enqueueIgnitionPendingAdd({
        ignitionClientSlug: clientMatch.id,
        clientName: updatedProfile.legalName || clientMatch.name,
        clientEmail: clientMatch.email,
        proposal,
        existingProfileId: updatedProfile.id,
        source: "webhook",
      });
    }
  } else {
    await enqueueIgnitionPendingAdd({
      ignitionClientSlug: clientMatch.id,
      clientName: clientMatch.name || proposal.clientName,
      clientEmail: clientMatch.email,
      proposal,
      source: "webhook",
    });
  }

  await writeIgnitionWebhooksStore({
    ...store,
    processedKeys: [dedupeKey, ...store.processedKeys].slice(0, 500),
  });

  return {
    handled: true,
    createdClients: [],
    updatedExisting,
  };
}

export function getWebhookSecretForVerification(
  store: IgnitionWebhooksStore,
): string {
  return resolveSigningSecret(store.subscription?.secret || "");
}

export type { IgnitionProposalRecord };
