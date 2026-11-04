import { acquireGraphAppOnlyToken } from "@/lib/azure-graph-app";
import {
  buildGraphMailHtml,
  parseEmailList,
  type GraphMailAttachment,
  type GraphMailPayload,
} from "@/lib/send-graph-mail";

/**
 * Unattended Graph send (client credentials / app-only).
 * Requires Application permission Mail.Send + admin consent,
 * and AZURE_CLIENT_SECRET (same app registration as SharePoint jobs).
 */

export type GraphAppMailPayload = {
  /** Mailbox UPN that appears as From (supervisor). */
  from: string;
  to: string;
  cc?: string;
  bcc?: string;
  subject: string;
  bodyText: string;
  highlights?: string[];
  portalHref?: string | null;
  signature: GraphMailPayload["signature"];
  attachments?: GraphMailAttachment[];
  /** Absolute origin for signature logo assets (e.g. https://clientportal…). */
  publicOrigin?: string;
};

function toRecipientList(emails: string[]) {
  return emails.map((address) => ({
    emailAddress: { address },
  }));
}

function withPublicOrigin(
  html: string,
  publicOrigin: string | undefined,
): string {
  const origin = (publicOrigin ?? "").replace(/\/+$/, "");
  if (!origin) {
    return html;
  }
  return html.replace(
    /(src=")(\/(?!\/)[^"]+")/gi,
    (_match, prefix: string, path: string) => `${prefix}${origin}${path}`,
  );
}

/**
 * Send mail as a specific user mailbox via Graph application permissions.
 */
export async function sendGraphAppMail(
  payload: GraphAppMailPayload,
): Promise<void> {
  const from = payload.from.trim();
  const to = parseEmailList(payload.to);
  if (!from) {
    throw new Error("From mailbox is required.");
  }
  if (to.length === 0) {
    throw new Error("Add at least one To email address.");
  }
  if (!payload.subject.trim()) {
    throw new Error("Subject is required.");
  }

  const token = await acquireGraphAppOnlyToken();
  let html = buildGraphMailHtml({
    to: payload.to,
    cc: payload.cc,
    bcc: payload.bcc,
    subject: payload.subject,
    bodyText: payload.bodyText,
    highlights: payload.highlights,
    portalHref: payload.portalHref,
    signature: payload.signature,
    attachments: payload.attachments,
  });
  html = withPublicOrigin(html, payload.publicOrigin);

  const message: Record<string, unknown> = {
    subject: payload.subject.trim(),
    body: {
      contentType: "HTML",
      content: html,
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

  const userPath = `/users/${encodeURIComponent(from)}`;
  const response = await fetch(
    `https://graph.microsoft.com/v1.0${userPath}/sendMail`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message,
        saveToSentItems: true,
      }),
    },
  );

  if (!response.ok) {
    let detail = "";
    try {
      const data = (await response.json()) as {
        error?: { message?: string; code?: string };
      };
      detail = data.error?.message?.trim() || data.error?.code || "";
    } catch {
      detail = "";
    }
    throw new Error(
      detail ||
        `Graph sendMail failed (${response.status}). Confirm Application Mail.Send + admin consent, and that the app may send as ${from}.`,
    );
  }
}
