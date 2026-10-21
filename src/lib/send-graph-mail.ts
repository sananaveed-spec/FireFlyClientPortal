import type {
  AccountInfo,
  IPublicClientApplication,
} from "@azure/msal-browser";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { mailSendRequest } from "@/auth/authConfig";
import {
  DEFAULT_SIGNATURE_LOGO,
  DEFAULT_SIGNATURE_WEBSITE,
  getSignatureLogoSrc,
  getSignatureWebsiteHref,
  hasEmailSignature,
  SIGNATURE_CONFIDENTIALITY_NOTICE,
  type EmailSignature,
} from "@/lib/email-signatures";
import { linkPortalKeywordsHtml } from "@/lib/email-portal-link";

export type GraphMailAttachment = {
  name: string;
  contentType: string;
  /** Base64-encoded file bytes for Microsoft Graph fileAttachment. */
  contentBytes: string;
};

export type GraphMailPayload = {
  to: string;
  cc?: string;
  bcc?: string;
  subject: string;
  bodyText: string;
  /** Terms to wrap in <strong> inside the body. */
  highlights?: string[];
  /** When set, whole-word "portal" in the body becomes a hyperlink. */
  portalHref?: string | null;
  signature: EmailSignature;
  /** Optional file attachments (e.g. BD8 question-list Excel). */
  attachments?: GraphMailAttachment[];
  /**
   * When set (must be in the future), Outlook holds the message until this
   * time via PidTagDeferredSendTime (SystemTime 0x3FEF).
   */
  sendAt?: Date;
};

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function parseEmailList(value: string): string[] {
  const seen = new Set<string>();
  const emails: string[] = [];
  for (const part of value.split(/[,;]+/)) {
    const email = part.trim();
    if (!email) {
      continue;
    }
    const key = email.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    emails.push(email);
  }
  return emails;
}

function toRecipientList(emails: string[]) {
  return emails.map((address) => ({
    emailAddress: { address },
  }));
}

function absoluteAssetUrl(src: string): string {
  const trimmed = src.trim();
  if (!trimmed) {
    return "";
  }
  if (/^https?:\/\//i.test(trimmed) || trimmed.startsWith("data:")) {
    return trimmed;
  }
  if (typeof window === "undefined") {
    return trimmed;
  }
  if (trimmed.startsWith("/")) {
    return `${window.location.origin}${trimmed}`;
  }
  return `${window.location.origin}/${trimmed}`;
}

function boldHighlightsHtml(text: string, highlights: string[]): string {
  const escaped = escapeHtml(text);
  const unique = [
    ...new Set(highlights.map((value) => value.trim()).filter(Boolean)),
  ].sort((a, b) => b.length - a.length);
  if (unique.length === 0) {
    return escaped;
  }
  const pattern = unique
    .map((value) => escapeHtml(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join("|");
  if (!pattern) {
    return escaped;
  }
  return escaped.replace(
    new RegExp(`(${pattern})`, "g"),
    '<strong style="font-weight:700;">$1</strong>',
  );
}

function buildSignatureHtml(signature: EmailSignature): string {
  if (!hasEmailSignature(signature)) {
    return "";
  }
  const lines: string[] = [];
  for (const value of [
    signature.name,
    signature.email,
    signature.position,
    signature.phone,
  ]) {
    const trimmed = value.trim();
    if (trimmed) {
      lines.push(
        `<div style="font-weight:700;color:#0a1628;">${escapeHtml(trimmed)}</div>`,
      );
    }
  }
  const website = signature.website.trim() || DEFAULT_SIGNATURE_WEBSITE;
  const websiteHref =
    getSignatureWebsiteHref(website) || DEFAULT_SIGNATURE_WEBSITE;
  lines.push(
    `<div><a href="${escapeHtml(websiteHref)}" style="color:#2f6fad;">${escapeHtml(website)}</a></div>`,
  );

  const logoSrc = absoluteAssetUrl(
    getSignatureLogoSrc(signature) || DEFAULT_SIGNATURE_LOGO,
  );
  if (logoSrc) {
    lines.push(
      `<div style="margin-top:8px;"><img src="${escapeHtml(logoSrc)}" alt="FireFly Accountancy" style="max-width:180px;height:auto;" /></div>`,
    );
  }

  lines.push(
    `<p style="margin-top:12px;font-size:11px;line-height:1.4;color:#555;">${escapeHtml(SIGNATURE_CONFIDENTIALITY_NOTICE)}</p>`,
  );

  return lines.join("\n");
}

export function buildGraphMailHtml(payload: GraphMailPayload): string {
  let bodyHtml = boldHighlightsHtml(
    payload.bodyText,
    payload.highlights ?? [],
  );
  const portalHref = payload.portalHref?.trim();
  if (portalHref) {
    bodyHtml = linkPortalKeywordsHtml(bodyHtml, portalHref, escapeHtml);
  }
  bodyHtml = bodyHtml.replace(/\r\n|\n|\r/g, "<br/>");
  const signatureHtml = buildSignatureHtml(payload.signature);

  return [
    `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.5;color:#0a1628;">`,
    bodyHtml,
    `<br/><br/>`,
    `<div>Thank you,</div>`,
    signatureHtml ? `<br/>${signatureHtml}` : "",
    `</div>`,
  ].join("");
}

async function acquireMailToken(
  instance: IPublicClientApplication,
  account: AccountInfo,
): Promise<string> {
  try {
    const result = await instance.acquireTokenSilent({
      ...mailSendRequest,
      account,
    });
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const result = await instance.acquireTokenPopup({
        ...mailSendRequest,
        account,
      });
      return result.accessToken;
    }
    throw error;
  }
}

/** Format a Date for `<input type="datetime-local">` (local wall clock). */
export function toDateTimeLocalValue(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return `${year}-${month}-${day}T${hour}:${minute}`;
}

export function parseDateTimeLocalValue(value: string): Date | null {
  const trimmed = value.trim();
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(trimmed)) {
    return null;
  }
  const parsed = new Date(trimmed);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }
  return parsed;
}

export async function sendGraphMail(
  instance: IPublicClientApplication,
  account: AccountInfo,
  payload: GraphMailPayload,
): Promise<void> {
  const to = parseEmailList(payload.to);
  if (to.length === 0) {
    throw new Error("Add at least one To email address.");
  }
  if (!payload.subject.trim()) {
    throw new Error("Subject is required.");
  }

  const sendAt = payload.sendAt;
  if (sendAt) {
    const soonest = Date.now() + 60_000;
    if (sendAt.getTime() < soonest) {
      throw new Error(
        "Choose a send time at least one minute in the future, or use Send Now.",
      );
    }
  }

  const token = await acquireMailToken(instance, account);
  const message: Record<string, unknown> = {
    subject: payload.subject.trim(),
    body: {
      contentType: "HTML",
      content: buildGraphMailHtml(payload),
    },
    toRecipients: toRecipientList(to),
    ccRecipients: toRecipientList(parseEmailList(payload.cc ?? "")),
    bccRecipients: toRecipientList(parseEmailList(payload.bcc ?? "")),
  };

  const attachments = (payload.attachments ?? []).filter(
    (item) => item.name.trim() && item.contentBytes.trim(),
  );
  if (attachments.length > 0) {
    message.hasAttachments = true;
    message.attachments = attachments.map((item) => ({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: item.name.trim(),
      contentType: item.contentType.trim() || "application/octet-stream",
      contentBytes: item.contentBytes.trim(),
    }));
  }

  if (sendAt) {
    message.singleValueExtendedProperties = [
      {
        id: "SystemTime 0x3FEF",
        value: sendAt.toISOString(),
      },
    ];
  }

  const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message,
      saveToSentItems: true,
    }),
  });

  if (!response.ok) {
    let detail = "";
    try {
      const data = (await response.json()) as {
        error?: { message?: string; code?: string };
      };
      detail = data.error?.message?.trim() || "";
    } catch {
      detail = "";
    }
    throw new Error(friendlyGraphMailError(detail, response.status));
  }
}

function friendlyGraphMailError(detail: string, status: number): string {
  const lower = detail.toLowerCase();
  if (
    lower.includes("inactive") ||
    lower.includes("soft-deleted") ||
    lower.includes("on-premise") ||
    lower.includes("on-premises") ||
    lower.includes("mailboxnotenabledforrestapiori")
  ) {
    return (
      "Your signed-in Microsoft account cannot send mail through Graph. " +
      "Use a Microsoft 365 mailbox with Exchange Online (not on-premises Exchange), " +
      "and make sure Mail.Send is granted for this app."
    );
  }
  if (lower.includes("mailboxnotenabledforrestapi") || status === 404) {
    return (
      "No cloud mailbox was found for your signed-in account. " +
      "Ask IT to assign an Exchange Online license, then sign out and sign in again."
    );
  }
  if (
    lower.includes("consent") ||
    lower.includes("access denied") ||
    status === 403
  ) {
    return (
      detail ||
      "Mail permission was denied. Ask an admin to grant Mail.Send for this app, then sign out and sign in again."
    );
  }
  return detail || `Failed to send email (${status}).`;
}
