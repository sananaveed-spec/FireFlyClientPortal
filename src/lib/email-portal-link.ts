import { sharePointPeriodFolder } from "@/lib/sharepoint-month-folders";

/** Marker left in body text for the single clickable "portal" word. */
export const PORTAL_WORD_MARKER = "⟦portal⟧";

function toAbsoluteHttpUrl(portalLink: string): string | null {
  const trimmed = portalLink.trim();
  if (!trimmed || /^\[portal link\]$/i.test(trimmed)) {
    return null;
  }
  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }
  return `https://${trimmed}`;
}

/**
 * Point the stored site portal link at an existing Shared Documents
 * year/month folder for the period (e.g. …/2026/10.Oct26). Does not create folders.
 */
export function buildSharePointMonthFolderHref(
  portalLink: string,
  forDate: Date = new Date(),
): string | null {
  const absolute = toAbsoluteHttpUrl(portalLink);
  if (!absolute) {
    return null;
  }

  let url: URL;
  try {
    url = new URL(absolute);
  } catch {
    return null;
  }

  if (!/\.sharepoint\.com$/i.test(url.hostname)) {
    return absolute;
  }

  const siteMatch = url.pathname.match(/^(\/sites\/[^/]+)/i);
  if (!siteMatch) {
    return absolute;
  }

  const sitePath = siteMatch[1];
  const { yearFolder, monthFolder } = sharePointPeriodFolder(forDate);
  const serverRelativePath = `${sitePath}/Shared Documents/${yearFolder}/${monthFolder}`;
  return `${url.origin}${sitePath}/Shared%20Documents/Forms/AllItems.aspx?id=${encodeURIComponent(serverRelativePath)}`;
}

/**
 * Resolve portal link to the period's existing SharePoint month folder.
 */
export function resolvePortalHref(
  portalLink: string,
  forDate: Date = new Date(),
): string | null {
  return buildSharePointMonthFolderHref(portalLink, forDate);
}

function tidyPortalText(text: string): string {
  return text.replace(/[ \t]+([.,;:])/g, "$1").replace(/[ \t]{2,}/g, " ");
}

/**
 * Resolve [portal] / [Portal Link] so exactly one "portal" word is marked for
 * linking (the upload instruction), not every mention of "portal" in the body.
 */
export function applyPortalLinkPlaceholders(
  text: string,
  portalLink: string,
  forDate: Date = new Date(),
): string {
  const href = resolvePortalHref(portalLink, forDate);
  let out = text;

  if (!href) {
    return tidyPortalText(out.replace(/\[portal\]/gi, "portal"));
  }

  // Explicit [portal] placeholder → one link marker
  if (/\[portal\]/i.test(out)) {
    out = out.replace(/\[portal\]/i, PORTAL_WORD_MARKER);
    out = out.replace(/\[portal\]/gi, "portal");
  }

  // Legacy ": [Portal Link]" / "[Portal Link]" — link "portal" on that line
  if (/\[Portal Link\]/i.test(out)) {
    out = out.replace(
      /^([^\n]*?)\bportal\b([^\n]*?)(?:\s*:\s*)?\[Portal Link\]/gim,
      (_match, before: string, after: string) =>
        `${before}${PORTAL_WORD_MARKER}${after}`,
    );
    if (/\[Portal Link\]/i.test(out) && !out.includes(PORTAL_WORD_MARKER)) {
      out = out.replace(/\[Portal Link\]/i, PORTAL_WORD_MARKER);
    }
    out = out
      .replace(/\s*:\s*\[Portal Link\]/gi, "")
      .replace(/\s*\[Portal Link\]/gi, "");
  }

  // Common upload line when templates already say plain "portal"
  if (!out.includes(PORTAL_WORD_MARKER)) {
    out = out.replace(
      /(Please upload (?:them|these|what you can) to the )portal\b/i,
      `$1${PORTAL_WORD_MARKER}`,
    );
  }
  if (!out.includes(PORTAL_WORD_MARKER)) {
    out = out.replace(
      /(upload (?:them|these|it) (?:directly against each item in|to) the )portal\b/i,
      `$1${PORTAL_WORD_MARKER}`,
    );
  }

  return tidyPortalText(out);
}

/** Turn the portal word marker into a blue website-style link (escaped body). */
export function linkPortalKeywordsHtml(
  escapedText: string,
  portalHref: string,
  escapeAttr: (value: string) => string,
): string {
  const href = portalHref.trim();
  if (!href || !escapedText || !escapedText.includes(PORTAL_WORD_MARKER)) {
    return escapedText;
  }
  const safeHref = escapeAttr(href);
  const anchor = `<a href="${safeHref}" style="color:#1a73e8;text-decoration:underline;">portal</a>`;
  return escapedText.split(PORTAL_WORD_MARKER).join(anchor);
}
