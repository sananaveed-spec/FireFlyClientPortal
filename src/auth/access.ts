import { isAllowedOrganizationEmail } from "@/auth/organization";
import {
  normalizeEmail,
  type PortalRole,
  type PortalUser,
} from "@/lib/portal-users";

export type ResolvedAccess = {
  email: string;
  role: PortalRole;
  canView: boolean;
  canEdit: boolean;
  canManageUsers: boolean;
};

export function canViewPortalDomain(email: string): boolean {
  return isAllowedOrganizationEmail(normalizeEmail(email));
}

export function resolveAccess(
  email: string,
  users: PortalUser[] | null,
): ResolvedAccess {
  const normalized = normalizeEmail(email);
  const onDomain = canViewPortalDomain(normalized);

  if (!onDomain) {
    return {
      email: normalized,
      role: "team",
      canView: false,
      canEdit: false,
      canManageUsers: false,
    };
  }

  const record = users?.find((user) => user.email === normalized) ?? null;

  if (!record) {
    return {
      email: normalized,
      role: "team",
      canView: false,
      canEdit: false,
      canManageUsers: false,
    };
  }

  if (record.role === "admin") {
    return {
      email: normalized,
      role: "admin",
      canView: true,
      canEdit: true,
      canManageUsers: true,
    };
  }

  return {
    email: normalized,
    role: "team",
    canView: true,
    canEdit: false,
    canManageUsers: false,
  };
}
