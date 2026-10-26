import { NextResponse } from "next/server";
import {
  dismissDashboardNotification,
  listActiveDashboardNotifications,
  readDashboardNotificationsStore,
} from "@/lib/dashboard-notifications";
import { requirePortalUser } from "@/lib/requireEditor";

export async function GET(request: Request) {
  const gate = await requirePortalUser(request);
  if (gate instanceof NextResponse) {
    return gate;
  }

  try {
    const store = await readDashboardNotificationsStore();
    return NextResponse.json({
      notifications: listActiveDashboardNotifications(store),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load dashboard notifications.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  const gate = await requirePortalUser(request);
  if (gate instanceof NextResponse) {
    return gate;
  }

  try {
    const body = (await request.json().catch(() => ({}))) as {
      action?: string;
      id?: string;
    };
    const action = String(body.action ?? "").trim().toLowerCase();
    const id = String(body.id ?? "").trim();

    if (action !== "dismiss" || !id) {
      return NextResponse.json(
        { error: 'Expected action "dismiss" with notification id.' },
        { status: 400 },
      );
    }

    const store = await dismissDashboardNotification(id);
    return NextResponse.json({
      notifications: listActiveDashboardNotifications(store),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to update dashboard notifications.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
