import { NextResponse } from "next/server";
import { requireAdmin } from "@/lib/requireEditor";
import {
  isPortalRole,
  isValidEmail,
  normalizeEmail,
} from "@/lib/portal-users";
import {
  readPortalUsers,
  removePortalUser,
  upsertPortalUser,
} from "@/lib/portal-users-store";

export async function GET() {
  try {
    const users = await readPortalUsers();
    return NextResponse.json({ users });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to load users.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json()) as {
      email?: unknown;
      role?: unknown;
      name?: unknown;
      position?: unknown;
      phone?: unknown;
    };
    const email =
      typeof body.email === "string" ? normalizeEmail(body.email) : "";
    const role = body.role;
    const name = typeof body.name === "string" ? body.name : "";
    const position = typeof body.position === "string" ? body.position : "";
    const phone = typeof body.phone === "string" ? body.phone : "";

    if (!isValidEmail(email)) {
      return NextResponse.json(
        { error: "Enter a valid email address." },
        { status: 400 },
      );
    }
    if (!isPortalRole(role)) {
      return NextResponse.json(
        { error: "Select a valid role (Admin or Team Member)." },
        { status: 400 },
      );
    }

    const result = await upsertPortalUser({
      email,
      role,
      name,
      position,
      phone,
    });
    if (result.error) {
      return NextResponse.json(
        { error: result.error, users: result.users },
        { status: 400 },
      );
    }

    return NextResponse.json({
      users: result.users,
      saved: result.saved,
      message: "User saved.",
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to save user.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function DELETE(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json()) as { email?: unknown };
    const email =
      typeof body.email === "string" ? normalizeEmail(body.email) : "";

    if (!isValidEmail(email)) {
      return NextResponse.json(
        { error: "Enter a valid email address." },
        { status: 400 },
      );
    }

    const result = await removePortalUser(email);
    if (result.error) {
      return NextResponse.json(
        { error: result.error, users: result.users },
        { status: 400 },
      );
    }

    return NextResponse.json({
      users: result.users,
      removed: result.removed,
      message: result.removed ? "User removed." : "User was not on the list.",
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to remove user.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
