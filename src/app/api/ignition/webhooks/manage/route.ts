import { NextResponse } from "next/server";
import {
  ensureIgnitionWebhookSubscription,
  getIgnitionWebhookReceiveUrl,
  listIgnitionWebhookSubscriptions,
  readIgnitionWebhooksStore,
} from "@/lib/ignition/webhooks";
import { requireAdmin } from "@/lib/requireEditor";

export async function GET(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const store = await readIgnitionWebhooksStore();
    let remote: Array<Record<string, unknown>> = [];
    let remoteError = "";
    try {
      remote = await listIgnitionWebhookSubscriptions();
    } catch (error) {
      remoteError =
        error instanceof Error ? error.message : "Failed to list remote webhooks.";
    }

    const receiveUrl = getIgnitionWebhookReceiveUrl();
    const matching = remote.filter(
      (item) => String(item.url ?? "").trim() === receiveUrl,
    );

    return NextResponse.json({
      receiveUrl,
      registered: Boolean(store.subscription?.slug) || matching.length > 0,
      hasSecret: Boolean(
        store.subscription?.secret ||
          process.env.IGNITION_WEBHOOK_SECRET?.trim(),
      ),
      subscription: store.subscription
        ? {
            slug: store.subscription.slug,
            url: store.subscription.url,
            events: store.subscription.events,
            status: store.subscription.status,
            createdAt: store.subscription.createdAt,
            updatedAt: store.subscription.updatedAt,
            // never return the raw secret to the browser
            hasSecret: Boolean(store.subscription.secret),
          }
        : null,
      remoteCount: matching.length,
      remoteError: remoteError || null,
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load Ignition webhook status.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

/**
 * Admin-only: register / refresh Ignition webhook_subscriptions for this server.
 */
export async function POST(request: Request) {
  const denied = await requireAdmin(request);
  if (denied) {
    return denied;
  }

  try {
    const body = (await request.json().catch(() => ({}))) as {
      action?: string;
    };
    const action = String(body.action ?? "ensure").trim().toLowerCase();
    if (action !== "ensure") {
      return NextResponse.json(
        { error: 'Only action "ensure" is supported.' },
        { status: 400 },
      );
    }

    const result = await ensureIgnitionWebhookSubscription();
    return NextResponse.json({
      message: result.message,
      created: result.created,
      receiveUrl: result.subscription.url,
      subscription: {
        slug: result.subscription.slug,
        url: result.subscription.url,
        events: result.subscription.events,
        status: result.subscription.status,
        createdAt: result.subscription.createdAt,
        updatedAt: result.subscription.updatedAt,
        hasSecret: Boolean(result.subscription.secret),
      },
      registered: true,
      hasSecret: true,
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to register Ignition webhook.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
