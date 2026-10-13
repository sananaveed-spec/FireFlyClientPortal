import { NextResponse } from "next/server";
import { resolveAccess } from "@/auth/access";
import { allowedEmailDomainsLabel } from "@/auth/organization";
import { normalizeEmail } from "@/lib/portal-users";
import { readPortalUsers } from "@/lib/portal-users-store";

/**
 * Soft server gate for write APIs: require x-user-email with Admin role.
 */
export async function requireEditor(
  request: Request,
): Promise<NextResponse | null> {
  const email = normalizeEmail(request.headers.get("x-user-email") ?? "");

  if (!email) {
    return NextResponse.json(
      {
        error: `Sign in with a ${allowedEmailDomainsLabel} account to continue.`,
      },
      { status: 401 },
    );
  }

  const users = await readPortalUsers();
  const access = resolveAccess(email, users);
  if (!access.canEdit) {
    return NextResponse.json(
      {
        error:
          "Admin access required to edit. Ask an Admin to update your role under User Management.",
      },
      { status: 403 },
    );
  }

  return null;
}

export const requireAdmin = requireEditor;

/**
 * Soft server gate: require a signed-in portal user (Admin or Team).
 */
export async function requirePortalUser(
  request: Request,
): Promise<NextResponse | { email: string }> {
  const email = normalizeEmail(request.headers.get("x-user-email") ?? "");

  if (!email) {
    return NextResponse.json(
      {
        error: `Sign in with a ${allowedEmailDomainsLabel} account to continue.`,
      },
      { status: 401 },
    );
  }

  const users = await readPortalUsers();
  const access = resolveAccess(email, users);
  if (!access.canView) {
    return NextResponse.json(
      {
        error:
          "Your account is not authorized for this portal. Ask an Admin to add you under User Management.",
      },
      { status: 403 },
    );
  }

  return { email };
}
