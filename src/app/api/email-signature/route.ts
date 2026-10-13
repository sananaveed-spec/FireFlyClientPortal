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
    const signature = await readEmailSignatureForUser(gate.email);
    return NextResponse.json({
      signature: await signatureWithPortalProfile(gate.email, signature),
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
