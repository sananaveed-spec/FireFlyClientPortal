import { NextResponse } from "next/server";
import {
  normalizeEmailSignature,
  type EmailSignature,
} from "@/lib/email-signatures";
import {
  readEmailSignatureForUser,
  writeEmailSignatureForUser,
} from "@/lib/data-store";
import { normalizeEmail } from "@/lib/portal-users";
import { readPortalUsers } from "@/lib/portal-users-store";
import { requirePortalUser } from "@/lib/requireEditor";

async function signatureWithPortalProfile(
  email: string,
  signature: EmailSignature,
): Promise<EmailSignature> {
  const users = await readPortalUsers();
  const portalUser = users.find(
    (user) => user.email === normalizeEmail(email),
  );
  return {
    ...signature,
    name: portalUser?.name?.trim() || signature.name,
    position: portalUser?.position?.trim() || signature.position,
    phone: portalUser?.phone?.trim() || signature.phone,
    email: email || signature.email,
  };
}

export async function GET(request: Request) {
  const gate = await requirePortalUser(request);
  if (gate instanceof NextResponse) {
    return gate;
  }

  try {
    // Optional ?email= loads that user's signature (e.g. supervisor for
    // Client Management schedule preview/send). Defaults to signed-in user.
    const requested = new URL(request.url).searchParams.get("email");
    const targetEmail = requested
      ? normalizeEmail(requested)
      : normalizeEmail(gate.email);
    if (!targetEmail) {
      return NextResponse.json(
        { error: "A valid email is required." },
        { status: 400 },
      );
    }
    const signature = await readEmailSignatureForUser(targetEmail);
    return NextResponse.json({
      signature: await signatureWithPortalProfile(targetEmail, signature),
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load email signature.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  const gate = await requirePortalUser(request);
  if (gate instanceof NextResponse) {
    return gate;
  }

  try {
    const body = (await request.json()) as { signature?: unknown };
    const signature = normalizeEmailSignature(body.signature);
    const saved = await writeEmailSignatureForUser(gate.email, signature);
    return NextResponse.json({ signature: saved } satisfies {
      signature: EmailSignature;
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to save email signature.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
