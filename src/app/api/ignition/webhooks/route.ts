import { NextResponse } from "next/server";
import {
  getWebhookSecretForVerification,
  processIgnitionWebhookEvent,
  readIgnitionWebhooksStore,
  verifyIgnitionWebhookSignature,
} from "@/lib/ignition/webhooks";

export const runtime = "nodejs";

/**
 * Ignition → portal webhook receiver (public HTTPS).
 * No portal login — authenticated via X-Ignition-Signature.
 */
export async function POST(request: Request) {
  const rawBody = await request.text();
  const signatureHeader =
    request.headers.get("x-ignition-signature") ||
    request.headers.get("X-Ignition-Signature") ||
    "";

  try {
    const store = await readIgnitionWebhooksStore();
    const secret = getWebhookSecretForVerification(store);
    if (!secret) {
      console.error(
        "Ignition webhook received but no signing secret is configured.",
      );
      return NextResponse.json(
        { error: "Webhook signing secret is not configured on the server." },
        { status: 503 },
      );
    }

    const valid = verifyIgnitionWebhookSignature({
      rawBody,
      signatureHeader,
      secret,
    });
    if (!valid) {
      return NextResponse.json(
        { error: "Invalid Ignition webhook signature." },
        { status: 401 },
      );
    }

    let parsed: unknown;
    try {
      parsed = JSON.parse(rawBody);
    } catch {
      return NextResponse.json({ error: "Invalid JSON body." }, { status: 400 });
    }

    const record =
      typeof parsed === "object" && parsed !== null
        ? (parsed as Record<string, unknown>)
        : {};
    const type = String(record.type ?? "").trim();
    const occurredAt = String(record.occurred_at ?? record.occurredAt ?? "").trim();
    const data = record.data;

    // Acknowledge quickly; processing is still awaited so retries are accurate.
    const result = await processIgnitionWebhookEvent({
      type,
      occurredAt,
      data,
    });

    return NextResponse.json({
      ok: true,
      ...result,
    });
  } catch (error) {
    console.error("Ignition webhook processing failed:", error);
    const message =
      error instanceof Error ? error.message : "Webhook processing failed.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function GET() {
  return NextResponse.json({
    ok: true,
    message:
      "Ignition webhook endpoint is live. POST proposal.accepted events here.",
  });
}
