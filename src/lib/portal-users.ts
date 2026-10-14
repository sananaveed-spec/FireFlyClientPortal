export const PORTAL_USERS_CHANGED_EVENT = "portal-users-changed";

export const PORTAL_ROLES = ["admin", "team"] as const;

export type PortalRole = (typeof PORTAL_ROLES)[number];

export type PortalUser = {
  email: string;
  role: PortalRole;
  name: string;
  position: string;
  phone: string;
};

/** Seed admins when the users file is empty or invalid. */
export const DEFAULT_ADMIN_EMAILS = [
  "sana@fireflyaccountancy.com",
  "sana.naveed@fireflyaccountancy.com",
  "sana@epsfresno.com",
] as const;

export const ROLE_LABELS: Record<PortalRole, string> = {
  admin: "Admin",
  team: "Team Member",
};

export function normalizeEmail(email: string): string {
  return email.trim().toLowerCase();
}

export function isValidEmail(email: string): boolean {
  const normalized = normalizeEmail(email);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized);
}

export function isPortalRole(value: unknown): value is PortalRole {
  return (
    typeof value === "string" &&
    (PORTAL_ROLES as readonly string[]).includes(value)
  );
}

export function sortPortalUsers(users: PortalUser[]): PortalUser[] {
  return [...users].sort((a, b) => a.email.localeCompare(b.email));
}

export function defaultAdminUsers(): PortalUser[] {
  return sortPortalUsers(
    DEFAULT_ADMIN_EMAILS.map((email) => ({
      email: normalizeEmail(email),
      role: "admin" as const,
      name: "",
      position: "",
      phone: "",
    })),
  );
}

export function parsePortalUsers(parsed: unknown): PortalUser[] {
  if (!Array.isArray(parsed)) {
    return defaultAdminUsers();
  }

  const byEmail = new Map<string, PortalUser>();

  for (const item of parsed) {
    if (typeof item === "string") {
      const email = normalizeEmail(item);
      if (!isValidEmail(email) || byEmail.has(email)) {
        continue;
      }
      byEmail.set(email, {
        email,
        role: "admin",
        name: "",
        position: "",
        phone: "",
      });
      continue;
    }

    if (!item || typeof item !== "object") {
      continue;
    }
    const record = item as Record<string, unknown>;
    const email =
      typeof record.email === "string" ? normalizeEmail(record.email) : "";
    if (!isValidEmail(email) || byEmail.has(email)) {
      continue;
    }

    const role: PortalRole = isPortalRole(record.role) ? record.role : "admin";
    byEmail.set(email, {
      email,
      role,
      name: String(record.name ?? "").trim(),
      position: String(record.position ?? "").trim(),
      phone: String(record.phone ?? "").trim(),
    });
  }

  const users = sortPortalUsers([...byEmail.values()]);
  return users.length > 0 ? users : defaultAdminUsers();
}

export function normalizePortalUser(input: {
  email: string;
  role: PortalRole;
  name?: string;
  position?: string;
  phone?: string;
}): PortalUser | { error: string } {
  const email = normalizeEmail(input.email);
  if (!isValidEmail(email)) {
    return { error: "Enter a valid email address." };
  }
  if (!isPortalRole(input.role)) {
    return { error: "Select a valid role." };
  }
  return {
    email,
    role: input.role,
    name: String(input.name ?? "").trim(),
    position: String(input.position ?? "").trim(),
    phone: String(input.phone ?? "").trim(),
  };
}

export function notifyPortalUsersChanged(): void {
  if (typeof window === "undefined") {
    return;
  }
  window.dispatchEvent(new Event(PORTAL_USERS_CHANGED_EVENT));
}
