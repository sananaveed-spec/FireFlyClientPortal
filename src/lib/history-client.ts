import { getAccountEmail } from "@/auth/organization";
import type { HistoryAction, HistoryEntityType, HistoryInput } from "@/lib/history";
import { portalFetch } from "@/lib/portal-api";

export type ActorInfo = {
  name: string;
  email: string;
};

export function getActorFromAccount(
  account?: { name?: string; username?: string } | null,
): ActorInfo {
  const email = getAccountEmail(account ?? undefined);
  return {
    name: account?.name?.trim() || email || "Unknown",
    email,
  };
}

export async function logHistoryEvent(input: HistoryInput) {
  const response = await portalFetch("/api/history", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });

  if (!response.ok) {
    const data = (await response.json()) as { error?: string };
    throw new Error(data.error || "Failed to write history.");
  }
}

export async function logPortalAction(params: {
  actor: ActorInfo;
  project: string;
  entityType: HistoryEntityType;
  entityLabel: string;
  action: HistoryAction;
  summary: string;
  details?: string;
  comment?: string;
}) {
  try {
    await logHistoryEvent({
      employeeName: params.actor.name,
      employeeEmail: params.actor.email,
      project: params.project,
      entityType: params.entityType,
      entityLabel: params.entityLabel,
      action: params.action,
      summary: params.summary,
      details: params.details ?? "",
      comment: params.comment ?? "",
    });
  } catch {
    // History must not block the main save/delete flow.
  }
}
