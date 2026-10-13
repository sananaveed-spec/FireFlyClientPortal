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

export type GraphMailPayload = {
  to: string;
  cc?: string;
  bcc?: string;
  subject: string;
  bodyText: string;
  /** Terms to wrap in <strong> inside the body. */
  highlights?: string[];
  signature: EmailSignature;
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
  const bodyHtml = boldHighlightsHtml(
    payload.bodyText,
    payload.highlights ?? [],
  ).replace(/\r\n|\n|\r/g, "<br/>");
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

  const token = await acquireMailToken(instance, account);
  const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: {
        subject: payload.subject.trim(),
        body: {
          contentType: "HTML",
          content: buildGraphMailHtml(payload),
        },
        toRecipients: toRecipientList(to),
        ccRecipients: toRecipientList(parseEmailList(payload.cc ?? "")),
        bccRecipients: toRecipientList(parseEmailList(payload.bcc ?? "")),
      },
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
