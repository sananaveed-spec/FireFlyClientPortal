export type EmailSignature = {
  name: string;
  position: string;
  phone: string;
  email: string;
  website: string;
  /**
   * Logo/image source. Defaults to the firm FireFly signature logo.
   * May be a public path or a data URL.
   */
  imageDataUrl: string;
  imageName: string;
  updatedAt: string;
  /** Legacy free-text signature (migrated into name when fields are empty). */
  text?: string;
};

export type EmailSignaturesStore = {
  signatures: Record<string, EmailSignature>;
};

/** Firm logo used automatically on every email signature. */
export const DEFAULT_SIGNATURE_LOGO = "/firefly-signature-logo.webp";

/** Firm website used on every email signature (clickable). */
export const DEFAULT_SIGNATURE_WEBSITE = "https://fireflyaccountancy.com/";

/** Fixed confidentiality notice appended to every email signature (not editable). */
export const SIGNATURE_CONFIDENTIALITY_NOTICE =
  "IMPORTANT: This email and any files transmitted with it are confidential and intended solely for the use of the individual or entity to whom they are addressed. Do not distribute or copy this email if you are not the named addressee. Please notify the sender immediately if received by mistake and delete this email. Disclosing, copying, distributing, or taking any action in reliance on the contents of this information is strictly prohibited.";

const MAX_IMAGE_CHARS = 2_500_000; // ~1.8MB binary as base64 data URL

export function createEmptyEmailSignature(): EmailSignature {
  return {
    name: "",
    position: "",
    phone: "",
    email: "",
    website: DEFAULT_SIGNATURE_WEBSITE,
    imageDataUrl: DEFAULT_SIGNATURE_LOGO,
    imageName: "FireFly Accountancy logo",
    updatedAt: "",
  };
}

function isDataImageUrl(value: string): boolean {
  return /^data:image\/(png|jpe?g|webp|gif);base64,/i.test(value);
}

function isAllowedImageSrc(value: string): boolean {
  if (!value) {
    return false;
  }
  if (value === DEFAULT_SIGNATURE_LOGO || value.startsWith("/")) {
    return true;
  }
  return isDataImageUrl(value) && value.length <= MAX_IMAGE_CHARS;
}

/** Always resolve a displayable logo (falls back to firm logo). */
export function getSignatureLogoSrc(signature: EmailSignature | null | undefined) {
  const src = signature?.imageDataUrl?.trim() ?? "";
  return isAllowedImageSrc(src) ? src : DEFAULT_SIGNATURE_LOGO;
}

function websiteHref(website: string): string {
  const trimmed = website.trim();
  if (!trimmed) {
    return "";
  }
  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }
  return `https://${trimmed}`;
}

/** Compose signature contact lines for display / copy. */
export function formatSignatureContactText(signature: EmailSignature): string {
  const contact = [
    signature.name.trim(),
    signature.position.trim(),
    signature.phone.trim(),
    signature.email.trim(),
    signature.website.trim(),
  ]
    .filter(Boolean)
    .join("\n");
  return [contact, SIGNATURE_CONFIDENTIALITY_NOTICE]
    .filter(Boolean)
    .join("\n\n");
}

export function normalizeEmailSignature(value: unknown): EmailSignature {
  const raw =
    typeof value === "object" && value !== null
      ? (value as Partial<EmailSignature> & { text?: unknown })
      : {};

  let name = String(raw.name ?? "").trim();
  let position = String(raw.position ?? "").trim();
  let phone = String(raw.phone ?? "").trim();
  let email = String(raw.email ?? "").trim();
  const legacyText = String(raw.text ?? "").trim();

  // Migrate older free-text signatures into Name when structured fields are empty.
  if (!name && !position && !phone && !email && legacyText) {
    const lines = legacyText
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    name = lines[0] ?? "";
    position = lines[1] ?? "";
    phone = lines[2] ?? "";
    email = lines[3] ?? "";
  }

  const rawImage = String(raw.imageDataUrl ?? "").trim();
  const imageDataUrl = isAllowedImageSrc(rawImage)
    ? rawImage
    : DEFAULT_SIGNATURE_LOGO;
  const imageName =
    String(raw.imageName ?? "").trim() ||
    (imageDataUrl === DEFAULT_SIGNATURE_LOGO
      ? "FireFly Accountancy logo"
      : "");
  const updatedAt = String(raw.updatedAt ?? "").trim();

  return {
    name,
    position,
    phone,
    email,
    website: DEFAULT_SIGNATURE_WEBSITE,
    imageDataUrl,
    imageName,
    updatedAt,
  };
}

export function normalizeEmailSignaturesStore(
  value: unknown,
): EmailSignaturesStore {
  const incoming =
    typeof value === "object" && value !== null
      ? (value as { signatures?: unknown })
      : {};
  const rawMap =
    typeof incoming.signatures === "object" && incoming.signatures !== null
      ? (incoming.signatures as Record<string, unknown>)
      : {};

  const signatures: Record<string, EmailSignature> = {};
  for (const [email, entry] of Object.entries(rawMap)) {
    const key = email.trim().toLowerCase();
    if (!key) {
      continue;
    }
    signatures[key] = normalizeEmailSignature(entry);
  }

  return { signatures };
}

export function hasEmailSignature(signature: EmailSignature | null | undefined) {
  if (!signature) {
    return false;
  }
  // Logo alone is not enough — require at least one contact field.
  return Boolean(
    signature.name.trim() ||
      signature.position.trim() ||
      signature.phone.trim() ||
      signature.email.trim() ||
      signature.website.trim(),
  );
}

export function getSignatureWebsiteHref(website: string): string {
  return websiteHref(website);
}

/** Hardcoded closing line shown before the email signature. */
export const SIGNATURE_REGARDS_LINE = "Regards,";

/**
 * Remove legacy closings from the end of a template body:
 * "Thank you," / "[Staff Name]" / "[Firm Name]" (whole lines only).
 * Leaves content like "Thank you for sending…" unchanged.
 */
export function stripLegacySignatureClosing(body: string): string {
  const lines = String(body ?? "")
    .replace(/\s+$/u, "")
    .split(/\r?\n/);

  while (lines.length > 0) {
    const last = (lines[lines.length - 1] ?? "").trim();
    if (
      !last ||
      /^thank you,?$/i.test(last) ||
      /^\[staff name\]$/i.test(last) ||
      /^\[firm name\]$/i.test(last)
    ) {
      lines.pop();
      continue;
    }
    break;
  }

  return lines.join("\n").replace(/\s+$/u, "");
}
