import { staffNameAliases } from "@/lib/client-management";
import {
  getLookupOptions,
  type LookupDataStore,
} from "@/lib/lookup-data";
import { normalizeEmail, type PortalUser } from "@/lib/portal-users";

/** Lookup field ids whose options come from User Management (not editable). */
export const USER_MANAGED_LOOKUP_FIELDS = [
  "primaryAssignee",
  "secondaryAssignee",
] as const;

export type UserManagedLookupFieldId =
  (typeof USER_MANAGED_LOOKUP_FIELDS)[number];

export type PortalUserOption = {
  value: string;
  label: string;
  meta?: string;
  keywords?: string;
};

export function isUserManagedLookupField(
  id: string,
): id is UserManagedLookupFieldId {
  return (USER_MANAGED_LOOKUP_FIELDS as readonly string[]).includes(id);
}

function toPortalUserOption(user: PortalUser): PortalUserOption {
  const name = user.name.trim();
  return {
    value: user.email,
    label: name || user.email,
    meta: name ? user.email : undefined,
    keywords: `${name} ${user.email}`,
  };
}

/** Dropdown options: value = email, label = name (email shown as meta). */
export function portalUserSelectOptions(
  users: PortalUser[],
): PortalUserOption[] {
  return [...users]
    .sort((a, b) => {
      const aLabel = (a.name.trim() || a.email).toLowerCase();
      const bLabel = (b.name.trim() || b.email).toLowerCase();
      return aLabel.localeCompare(bLabel);
    })
    .map(toPortalUserOption);
}

/**
 * Resource Primary/Secondary options = Lookup curated list ∩ User Management.
 * Lookup item labels are emails (preferred) or legacy names.
 */
export function assigneeOptionsFromLookup(
  store: LookupDataStore | null | undefined,
  fieldId: UserManagedLookupFieldId,
  users: PortalUser[],
): PortalUserOption[] {
  const labels = getLookupOptions(store, fieldId);
  const options: PortalUserOption[] = [];
  const seen = new Set<string>();

  for (const label of labels) {
    const user = findPortalUserByStoredLabel(label, users);
    if (user) {
      if (seen.has(user.email)) {
        continue;
      }
      seen.add(user.email);
      options.push(toPortalUserOption(user));
      continue;
    }
    const trimmed = label.trim();
    if (!trimmed || seen.has(normalizeEmail(trimmed))) {
      continue;
    }
    const value = trimmed.includes("@") ? normalizeEmail(trimmed) : trimmed;
    seen.add(value);
    options.push({
      value,
      label: trimmed,
      keywords: trimmed,
    });
  }

  return options.sort((a, b) =>
    a.label.localeCompare(b.label, undefined, { sensitivity: "base" }),
  );
}

export function findPortalUserByStoredLabel(
  stored: string,
  users: PortalUser[],
): PortalUser | null {
  const trimmed = stored.trim();
  if (!trimmed) {
    return null;
  }
  const asEmail = normalizeEmail(trimmed);
  if (trimmed.includes("@")) {
    return users.find((user) => user.email === asEmail) ?? null;
  }
  const lower = trimmed.toLowerCase().replace(/\s+/g, " ");
  const exact = users.find(
    (user) => user.name.trim().toLowerCase().replace(/\s+/g, " ") === lower,
  );
  if (exact) {
    return exact;
  }
  const aliases = staffNameAliases(trimmed);
  return (
    users.find((user) => {
      const nameAliases = staffNameAliases(user.name);
      return aliases.some((alias) => nameAliases.includes(alias));
    }) ?? null
  );
}

/** Prefer email for storage; keep legacy name if user not found. */
export function coerceStaffToEmail(
  stored: string,
  users: PortalUser[],
): string {
  const trimmed = stored.trim();
  if (!trimmed) {
    return "";
  }
  const user = findPortalUserByStoredLabel(trimmed, users);
  return user?.email ?? (trimmed.includes("@") ? normalizeEmail(trimmed) : trimmed);
}

export function resolveStaffEmail(
  stored: string,
  users: PortalUser[],
): string {
  const user = findPortalUserByStoredLabel(stored, users);
  if (user) {
    return user.email;
  }
  const trimmed = stored.trim();
  return trimmed.includes("@") ? normalizeEmail(trimmed) : "";
}

/** Display name (email) for tables and From hints. */
export function formatStaffAssignmentLabel(
  stored: string,
  users: PortalUser[],
): string {
  const trimmed = stored.trim();
  if (!trimmed) {
    return "";
  }
  const user = findPortalUserByStoredLabel(trimmed, users);
  if (!user) {
    return trimmed;
  }
  const name = user.name.trim();
  return name ? `${name} (${user.email})` : user.email;
}

export function formatStaffAssignmentLabels(
  values: string[],
  users: PortalUser[],
): string {
  return values
    .map((value) => formatStaffAssignmentLabel(value, users))
    .filter(Boolean)
    .join(" + ");
}

/** Match keys for assignment: email + name aliases. */
export function userMatchKeys(
  email: string,
  ...names: Array<string | null | undefined>
): string[] {
  const keys = new Set<string>();
  const normalized = normalizeEmail(email);
  if (normalized) {
    keys.add(normalized);
  }
  for (const alias of staffNameAliases(...names)) {
    keys.add(alias);
  }
  return [...keys];
}
