import {
  defaultAdminUsers,
  normalizeEmail,
  normalizePortalUser,
  parsePortalUsers,
  type PortalRole,
  type PortalUser,
} from "@/lib/portal-users";
import { readPortalUsersFile, writePortalUsersFile } from "@/lib/data-store";

export async function readPortalUsers(): Promise<PortalUser[]> {
  const parsed = await readPortalUsersFile();
  return parsePortalUsers(parsed);
}

export async function writePortalUsers(
  users: PortalUser[],
): Promise<PortalUser[]> {
  const next = parsePortalUsers(users);
  const toWrite = next.length > 0 ? next : defaultAdminUsers();
  await writePortalUsersFile(toWrite);
  return toWrite;
}

function adminCount(users: PortalUser[]): number {
  return users.filter((user) => user.role === "admin").length;
}

export async function upsertPortalUser(input: {
  email: string;
  role: PortalRole;
  name?: string;
  position?: string;
  phone?: string;
}): Promise<{ users: PortalUser[]; saved: boolean; error?: string }> {
  const current = await readPortalUsers();
  const existing = current.find(
    (user) => user.email === normalizeEmail(input.email),
  );

  const normalized = normalizePortalUser({
    email: input.email,
    role: input.role,
    name: input.name !== undefined ? input.name : (existing?.name ?? ""),
    position:
      input.position !== undefined
        ? input.position
        : (existing?.position ?? ""),
    phone: input.phone !== undefined ? input.phone : (existing?.phone ?? ""),
  });
  if ("error" in normalized) {
    return { users: current, saved: false, error: normalized.error };
  }

  const without = current.filter((user) => user.email !== normalized.email);

  if (
    existing?.role === "admin" &&
    normalized.role !== "admin" &&
    adminCount(current) <= 1
  ) {
    return {
      users: current,
      saved: false,
      error: "At least one admin is required.",
    };
  }

  const users = await writePortalUsers([...without, normalized]);
  return { users, saved: true };
}

export async function removePortalUser(email: string): Promise<{
  users: PortalUser[];
  removed: boolean;
  error?: string;
}> {
  const normalized = normalizeEmail(email);
  const current = await readPortalUsers();
  const existing = current.find((user) => user.email === normalized);

  if (!existing) {
    return { users: current, removed: false };
  }

  if (existing.role === "admin" && adminCount(current) <= 1) {
    return {
      users: current,
      removed: false,
      error: "At least one admin is required.",
    };
  }

  const users = await writePortalUsers(
    current.filter((user) => user.email !== normalized),
  );
  return { users, removed: true };
}
