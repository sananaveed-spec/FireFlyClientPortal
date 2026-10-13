import type { ClientProfileEntry } from "@/lib/client-profiles";
import type { ResourceClientEntry } from "@/lib/resource-clients";

export type ClientAssignmentRole = "supervisor" | "bookkeeper";

export type AssignedClient = {
  resource: ResourceClientEntry;
  profile: ClientProfileEntry | null;
  roles: ClientAssignmentRole[];
};

/** Build match aliases from a person's display name (full + first name). */
export function staffNameAliases(...names: Array<string | null | undefined>): string[] {
  const aliases = new Set<string>();
  for (const name of names) {
    const normalized = String(name ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
    if (!normalized) {
      continue;
    }
    aliases.add(normalized);
    const first = normalized.split(" ")[0];
    if (first) {
      aliases.add(first);
    }
  }
  return [...aliases];
}

function staffLabelMatches(
  staffLabel: string,
  aliases: string[],
): boolean {
  const label = staffLabel.trim().toLowerCase().replace(/\s+/g, " ");
  if (!label || aliases.length === 0) {
    return false;
  }
  if (aliases.includes(label)) {
    return true;
  }
  const labelFirst = label.split(" ")[0] ?? "";
  return Boolean(labelFirst && aliases.includes(labelFirst));
}

export function rolesForUserOnClient(
  client: ResourceClientEntry,
  aliases: string[],
): ClientAssignmentRole[] {
  const roles: ClientAssignmentRole[] = [];
  if (staffLabelMatches(client.supervisor, aliases)) {
    roles.push("supervisor");
  }
  if (
    client.bookkeeperAccountant.some((entry) =>
      staffLabelMatches(entry, aliases),
    )
  ) {
    roles.push("bookkeeper");
  }
  return roles;
}

function normalizeClientNameKey(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[.,]/g, " ")
    .replace(
      /\b(incorporated|inc|llc|ltd|limited|corp|corporation|co)\b/g,
      " ",
    )
    .replace(/\s+/g, " ")
    .trim();
}

export function findProfileForResourceClient(
  client: ResourceClientEntry,
  profiles: ClientProfileEntry[],
): ClientProfileEntry | null {
  const key = client.clientName.trim().toLowerCase();
  if (!key) {
    return null;
  }

  const exact =
    profiles.find(
      (profile) => profile.legalName.trim().toLowerCase() === key,
    ) ??
    profiles.find(
      (profile) => profile.operatingName.trim().toLowerCase() === key,
    );
  if (exact) {
    return exact;
  }

  const normalizedKey = normalizeClientNameKey(client.clientName);
  if (!normalizedKey) {
    return null;
  }

  return (
    profiles.find((profile) => {
      const legal = normalizeClientNameKey(profile.legalName);
      const operating = normalizeClientNameKey(profile.operatingName);
      return (
        legal === normalizedKey ||
        operating === normalizedKey ||
        legal.startsWith(normalizedKey) ||
        normalizedKey.startsWith(legal) ||
        (operating !== "" &&
          (operating.startsWith(normalizedKey) ||
            normalizedKey.startsWith(operating)))
      );
    }) ?? null
  );
}

/** Clients where the user is Supervisor or Bookkeeper/Accountant. */
export function getClientsAssignedToUser(
  resources: ResourceClientEntry[],
  profiles: ClientProfileEntry[],
  aliases: string[],
): AssignedClient[] {
  if (aliases.length === 0) {
    return [];
  }

  const assigned: AssignedClient[] = [];
  for (const resource of resources) {
    const roles = rolesForUserOnClient(resource, aliases);
    if (roles.length === 0) {
      continue;
    }
    assigned.push({
      resource,
      profile: findProfileForResourceClient(resource, profiles),
      roles,
    });
  }

  return assigned.sort((a, b) =>
    a.resource.clientName.localeCompare(b.resource.clientName, undefined, {
      sensitivity: "base",
    }),
  );
}

export const ASSIGNMENT_ROLE_LABELS: Record<ClientAssignmentRole, string> = {
  supervisor: "Supervisor",
  bookkeeper: "Bookkeeper/Accountant",
};
