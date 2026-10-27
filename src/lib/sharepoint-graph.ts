import type {
  AccountInfo,
  IPublicClientApplication,
} from "@azure/msal-browser";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { sharePointRequest } from "@/auth/authConfig";
import {
  sharePointPeriodFolder,
  sharePointYearMonthFolderNames,
} from "@/lib/sharepoint-month-folders";

export type SharePointSiteOption = {
  id: string;
  displayName: string;
  webUrl: string;
};

const GRAPH = "https://graph.microsoft.com/v1.0";

/** Hostname from NEXT_PUBLIC_SHAREPOINT_SITE_HOST (default Fire Fly). */
export function sharePointSiteHost(): string {
  return (
    process.env.NEXT_PUBLIC_SHAREPOINT_SITE_HOST?.trim() ||
    "fireflyaccountancy.sharepoint.com"
  );
}

export function clientSharePointSiteDisplayName(legalName: string): string {
  const name = legalName.trim();
  return name ? `CLIENT - ${name}` : "CLIENT - New Client";
}

function toMailNickname(legalName: string): string {
  const cleaned = legalName.replace(/[^a-zA-Z0-9]/g, "");
  const base = (cleaned || "Client").slice(0, 48);
  return `CLIENT${base}`.slice(0, 64);
}

async function acquireSharePointToken(
  instance: IPublicClientApplication,
  account: AccountInfo,
): Promise<string> {
  try {
    const result = await instance.acquireTokenSilent({
      ...sharePointRequest,
      account,
    });
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const result = await instance.acquireTokenPopup({
        ...sharePointRequest,
        account,
      });
      return result.accessToken;
    }
    throw error;
  }
}

async function graphJson<T>(
  accessToken: string,
  url: string,
  init?: RequestInit,
): Promise<T> {
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      ...(init?.body ? { "Content-Type": "application/json" } : {}),
      ...init?.headers,
    },
  });
  const payload = (await response.json().catch(() => ({}))) as T & {
    error?: { message?: string; code?: string };
  };
  if (!response.ok) {
    const message =
      payload.error?.message ||
      `SharePoint Graph request failed (${response.status}).`;
    if (/invalid hostname for this tenancy/i.test(message)) {
      throw new Error(
        "SharePoint is on the Fire Fly tenant. Sign in with a @fireflyaccountancy.com account (not sana@epsfresno.com), then try again.",
      );
    }
    throw new Error(message);
  }
  return payload;
}

type GraphGroup = {
  id?: string;
  displayName?: string;
  groupTypes?: string[];
};

async function listClientGroups(
  accessToken: string,
): Promise<GraphGroup[]> {
  const byId = new Map<string, GraphGroup>();

  async function collect(url: string, headers?: HeadersInit) {
    try {
      const data = await graphJson<{ value?: GraphGroup[] }>(
        accessToken,
        url,
        headers ? { headers } : undefined,
      );
      for (const group of data.value ?? []) {
        const id = String(group.id ?? "").trim();
        const displayName = String(group.displayName ?? "").trim();
        if (!id || !displayName) {
          continue;
        }
        if (!isClientSharePointSite({ displayName, webUrl: "" })) {
          continue;
        }
        byId.set(id, group);
      }
    } catch {
      // Optional source — ignore failures.
    }
  }

  await collect(
    `${GRAPH}/me/memberOf/microsoft.graph.group?$select=id,displayName,groupTypes&$top=999`,
  );
  await collect(
    `${GRAPH}/me/ownedObjects/microsoft.graph.group?$select=id,displayName,groupTypes&$top=999`,
  );
  await collect(
    `${GRAPH}/groups?$search="displayName:CLIENT"&$select=id,displayName,groupTypes&$count=true&$top=100`,
    { ConsistencyLevel: "eventual" },
  );
  await collect(
    `${GRAPH}/groups?$search="displayName:CLIENTS"&$select=id,displayName,groupTypes&$count=true&$top=100`,
    { ConsistencyLevel: "eventual" },
  );

  return Array.from(byId.values());
}

async function siteFromGroup(
  accessToken: string,
  group: GraphGroup,
): Promise<SharePointSiteOption | null> {
  const groupId = String(group.id ?? "").trim();
  const displayName = String(group.displayName ?? "").trim();
  if (!groupId) {
    return null;
  }
  try {
    const site = await graphJson<{ id?: string; webUrl?: string; displayName?: string }>(
      accessToken,
      `${GRAPH}/groups/${groupId}/sites/root?$select=id,webUrl,displayName`,
    );
    const webUrl = String(site.webUrl ?? "").trim();
    if (!webUrl) {
      return null;
    }
    return {
      id: String(site.id ?? webUrl),
      displayName: String(site.displayName ?? displayName).trim() || displayName,
      webUrl,
    };
  } catch {
    return null;
  }
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function mergeSiteOptions(
  into: Map<string, SharePointSiteOption>,
  sites: SharePointSiteOption[],
) {
  for (const site of sites) {
    const key = site.webUrl.toLowerCase();
    if (!into.has(key)) {
      into.set(key, site);
    }
  }
}

/** Client portal sites only: "CLIENT - …" / "CLIENTS - …" (not client-portal). */
export function isClientSharePointSite(site: {
  displayName: string;
  webUrl: string;
}): boolean {
  // Require "CLIENT - " or "CLIENTS - " with spaces around the dash.
  return /^clients?\s+-\s+\S/i.test(site.displayName.trim());
}

/** List SharePoint sites for Portal Link picker (needs Sites.Read.All). */
export async function listSharePointSites(
  instance: IPublicClientApplication,
  account: AccountInfo,
  search = "*",
): Promise<SharePointSiteOption[]> {
  const accessToken = await acquireSharePointToken(instance, account);
  const query = encodeURIComponent(search.trim() || "*");
  const data = await graphJson<{
    value?: { id?: string; displayName?: string; webUrl?: string }[];
  }>(
    accessToken,
    `${GRAPH}/sites?search=${query}&$select=id,displayName,webUrl&$top=100`,
  );

  const byUrl = new Map<string, SharePointSiteOption>();
  for (const site of data.value ?? []) {
    const webUrl = String(site.webUrl ?? "").trim();
    const displayName = String(site.displayName ?? "").trim();
    if (!webUrl || !/sharepoint\.com/i.test(webUrl)) {
      continue;
    }
    byUrl.set(webUrl.toLowerCase(), {
      id: String(site.id ?? webUrl),
      displayName: displayName || webUrl,
      webUrl,
    });
  }
  return Array.from(byUrl.values()).sort((a, b) =>
    a.displayName.localeCompare(b.displayName, undefined, {
      sensitivity: "base",
    }),
  );
}

/**
 * Load CLIENT -/CLIENTS - sites for the Portal Link picker.
 * Site search alone often misses group-connected sites; also resolve from
 * groups the user owns/belongs to (e.g. CLIENTS - The Pantry).
 */
export async function listSharePointSitesForPicker(
  instance: IPublicClientApplication,
  account: AccountInfo,
  legalName = "",
): Promise<SharePointSiteOption[]> {
  const queries = ["CLIENT", "CLIENTS", legalName.trim()].filter(
    (value, index, all) => value && all.indexOf(value) === index,
  );

  const byUrl = new Map<string, SharePointSiteOption>();
  for (const query of queries) {
    try {
      mergeSiteOptions(
        byUrl,
        await listSharePointSites(instance, account, query),
      );
    } catch {
      // Keep going; another query may still succeed.
    }
  }

  try {
    const accessToken = await acquireSharePointToken(instance, account);
    const groups = await listClientGroups(accessToken);
    for (const group of groups) {
      const site = await siteFromGroup(accessToken, group);
      if (site) {
        mergeSiteOptions(byUrl, [site]);
      }
    }
  } catch {
    // Group lookup is best-effort on top of site search.
  }

  return Array.from(byUrl.values())
    .filter(isClientSharePointSite)
    .sort((a, b) =>
      a.displayName.localeCompare(b.displayName, undefined, {
        sensitivity: "base",
      }),
    );
}

/** Parse comma/semicolon/newline-separated emails. */
export function parseMemberEmails(value: string): string[] {
  const seen = new Set<string>();
  const emails: string[] = [];
  for (const part of value.split(/[,;\n]+/)) {
    const email = part.trim();
    if (!email || !email.includes("@")) {
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

async function resolveDirectoryUserId(
  accessToken: string,
  email: string,
): Promise<string> {
  const trimmed = email.trim();
  const encoded = encodeURIComponent(trimmed);
  try {
    const user = await graphJson<{ id?: string }>(
      accessToken,
      `${GRAPH}/users/${encoded}?$select=id`,
    );
    if (user.id) {
      return user.id;
    }
  } catch {
    // Fall through to mail/UPN filter.
  }

  const safe = trimmed.replace(/'/g, "''");
  const filter = encodeURIComponent(
    `mail eq '${safe}' or userPrincipalName eq '${safe}'`,
  );
  const data = await graphJson<{ value?: { id?: string }[] }>(
    accessToken,
    `${GRAPH}/users?$filter=${filter}&$select=id&$top=1`,
  );
  const id = data.value?.[0]?.id?.trim();
  if (!id) {
    throw new Error("User not found in the Fire Fly directory.");
  }
  return id;
}

async function addGroupMember(
  accessToken: string,
  groupId: string,
  email: string,
): Promise<void> {
  const userId = await resolveDirectoryUserId(accessToken, email);
  await graphJson(accessToken, `${GRAPH}/groups/${groupId}/members/$ref`, {
    method: "POST",
    body: JSON.stringify({
      "@odata.id": `${GRAPH}/directoryObjects/${userId}`,
    }),
  });
}

function currentYearMonthFolderNames(now = new Date()): {
  yearFolder: string;
  monthFolders: string[];
} {
  return sharePointYearMonthFolderNames(now.getFullYear());
}

async function listChildFolderNames(
  accessToken: string,
  childrenUrl: string,
): Promise<Map<string, string>> {
  const data = await graphJson<{
    value?: { id?: string; name?: string; folder?: unknown }[];
  }>(accessToken, childrenUrl);
  const byName = new Map<string, string>();
  for (const item of data.value ?? []) {
    if (!item.folder || !item.id || !item.name) {
      continue;
    }
    byName.set(item.name, item.id);
  }
  return byName;
}

async function ensureChildFolder(
  accessToken: string,
  parentChildrenUrl: string,
  createUrl: string,
  folderName: string,
  existing: Map<string, string>,
): Promise<string> {
  const already = existing.get(folderName);
  if (already) {
    return already;
  }
  try {
    const created = await graphJson<{ id?: string }>(accessToken, createUrl, {
      method: "POST",
      body: JSON.stringify({
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      }),
    });
    if (created.id) {
      existing.set(folderName, created.id);
      return created.id;
    }
  } catch {
    // Likely already exists — re-list and use it.
  }
  const refreshed = await listChildFolderNames(accessToken, parentChildrenUrl);
  for (const [name, id] of refreshed) {
    existing.set(name, id);
  }
  const id = existing.get(folderName);
  if (!id) {
    throw new Error(`Could not create or find folder "${folderName}".`);
  }
  return id;
}

/** Create Shared Documents/{year}/JanYY…DecYY if missing. */
async function ensureYearMonthFolders(
  accessToken: string,
  groupId: string,
): Promise<string[]> {
  const warnings: string[] = [];
  const { yearFolder, monthFolders } = currentYearMonthFolderNames();

  let driveId = "";
  for (let attempt = 0; attempt < 8; attempt += 1) {
    try {
      const drive = await graphJson<{ id?: string }>(
        accessToken,
        `${GRAPH}/groups/${groupId}/sites/root/drive?$select=id`,
      );
      driveId = String(drive.id ?? "").trim();
      if (driveId) {
        break;
      }
    } catch {
      // Drive may not be ready immediately after site create.
    }
    await sleep(2000);
  }
  if (!driveId) {
    return ["Could not open Shared Documents to create year/month folders."];
  }

  try {
    const rootChildrenUrl = `${GRAPH}/drives/${driveId}/root/children?$select=id,name,folder&$top=200`;
    const rootFolders = await listChildFolderNames(
      accessToken,
      rootChildrenUrl,
    );
    const yearId = await ensureChildFolder(
      accessToken,
      rootChildrenUrl,
      `${GRAPH}/drives/${driveId}/root/children`,
      yearFolder,
      rootFolders,
    );

    const yearChildrenUrl = `${GRAPH}/drives/${driveId}/items/${yearId}/children?$select=id,name,folder&$top=200`;
    const monthExisting = await listChildFolderNames(
      accessToken,
      yearChildrenUrl,
    );
    for (const monthName of monthFolders) {
      try {
        await ensureChildFolder(
          accessToken,
          yearChildrenUrl,
          `${GRAPH}/drives/${driveId}/items/${yearId}/children`,
          monthName,
          monthExisting,
        );
      } catch (monthError) {
        const detail =
          monthError instanceof Error
            ? monthError.message
            : "Could not create month folder.";
        warnings.push(`${monthName}: ${detail}`);
      }
    }
  } catch (folderError) {
    warnings.push(
      folderError instanceof Error
        ? folderError.message
        : "Failed to create year/month folders.",
    );
  }

  return warnings;
}

/**
 * Create a private Microsoft 365 Group (Team site) named CLIENT - {Legal Name}.
 * Needs Group.ReadWrite.All. Returns the SharePoint site webUrl.
 * Optional memberEmails are added as group members (not owners).
 * Also creates Shared Documents/{year}/JanYY…DecYY (skips existing).
 */
export async function createClientSharePointSite(
  instance: IPublicClientApplication,
  account: AccountInfo,
  legalName: string,
  memberEmails: string[] = [],
): Promise<{
  displayName: string;
  webUrl: string;
  memberWarnings: string[];
}> {
  const name = legalName.trim();
  if (!name) {
    throw new Error("Legal Name is required to create a SharePoint site.");
  }

  const displayName = clientSharePointSiteDisplayName(name);
  const accessToken = await acquireSharePointToken(instance, account);
  const mailNickname = toMailNickname(name);
  const creatorEmail = String(
    account.username || account.name || "",
  ).toLowerCase();

  const group = await graphJson<{ id: string }>(accessToken, `${GRAPH}/groups`, {
    method: "POST",
    body: JSON.stringify({
      displayName,
      description: `Client portal site for ${name}`,
      groupTypes: ["Unified"],
      mailEnabled: true,
      mailNickname,
      securityEnabled: false,
      visibility: "Private",
    }),
  });

  const memberWarnings: string[] = [];
  for (const email of memberEmails) {
    if (email.toLowerCase() === creatorEmail) {
      continue; // Creator is already owner/member.
    }
    try {
      await addGroupMember(accessToken, group.id, email);
    } catch (memberError) {
      const detail =
        memberError instanceof Error
          ? memberError.message
          : "Could not add member.";
      memberWarnings.push(`${email}: ${detail}`);
    }
  }

  let webUrl = "";
  let resolvedDisplayName = displayName;
  for (let attempt = 0; attempt < 12; attempt += 1) {
    try {
      const site = await graphJson<{ webUrl?: string; displayName?: string }>(
        accessToken,
        `${GRAPH}/groups/${group.id}/sites/root`,
      );
      webUrl = String(site.webUrl ?? "").trim();
      if (webUrl) {
        resolvedDisplayName = String(site.displayName ?? displayName).trim();
        break;
      }
    } catch {
      // Site provisioning can lag behind group creation.
    }
    await sleep(2500);
  }

  if (!webUrl) {
    const host = sharePointSiteHost();
    webUrl = `https://${host}/sites/${mailNickname}`;
  }

  const folderWarnings = await ensureYearMonthFolders(accessToken, group.id);
  memberWarnings.push(...folderWarnings);

  return {
    displayName: resolvedDisplayName,
    webUrl,
    memberWarnings,
  };
}

function normalizeSharePointSiteKey(url: string): string {
  try {
    const parsed = new URL(
      /^https?:\/\//i.test(url) ? url : `https://${url}`,
    );
    const match = parsed.pathname.match(/^(\/sites\/[^/]+)/i);
    const path = match?.[1] ?? "";
    return `${parsed.hostname.toLowerCase()}${path.toLowerCase()}`;
  } catch {
    return url.trim().toLowerCase().replace(/\/+$/, "");
  }
}

/**
 * Resolve the Microsoft 365 group id for a client portal SharePoint URL
 * (sites created via createClientSharePointSite are group-backed).
 */
async function resolveGroupIdFromPortalLink(
  accessToken: string,
  portalLink: string,
): Promise<string> {
  const site = parseSharePointSiteFromPortalLink(portalLink);
  if (!site) {
    throw new Error(
      "Portal Link must be a SharePoint site URL (e.g. https://….sharepoint.com/sites/…).",
    );
  }

  const siteName = site.sitePath.replace(/^\/sites\//i, "").trim();
  if (siteName) {
    try {
      const safe = siteName.replace(/'/g, "''");
      const data = await graphJson<{ value?: { id?: string }[] }>(
        accessToken,
        `${GRAPH}/groups?$filter=mailNickname eq '${safe}'&$select=id&$top=1`,
      );
      const id = String(data.value?.[0]?.id ?? "").trim();
      if (id) {
        return id;
      }
    } catch {
      // Fall through to site / group matching.
    }
  }

  const siteLookupPath = `${site.hostname}:${site.sitePath}`;
  const siteInfo = await graphJson<{
    id?: string;
    displayName?: string;
    webUrl?: string;
  }>(
    accessToken,
    `${GRAPH}/sites/${siteLookupPath}?$select=id,displayName,webUrl`,
  );
  const targetKey = normalizeSharePointSiteKey(
    String(siteInfo.webUrl ?? "").trim() || portalLink,
  );
  const displayName = String(siteInfo.displayName ?? "").trim();

  async function groupIdIfSiteMatches(group: GraphGroup): Promise<string | null> {
    const groupId = String(group.id ?? "").trim();
    if (!groupId) {
      return null;
    }
    const rooted = await siteFromGroup(accessToken, group);
    if (!rooted) {
      return null;
    }
    if (normalizeSharePointSiteKey(rooted.webUrl) !== targetKey) {
      return null;
    }
    return groupId;
  }

  if (displayName) {
    try {
      const search = encodeURIComponent(`"displayName:${displayName}"`);
      const data = await graphJson<{ value?: GraphGroup[] }>(
        accessToken,
        `${GRAPH}/groups?$search=${search}&$select=id,displayName,groupTypes&$count=true&$top=25`,
        { headers: { ConsistencyLevel: "eventual" } },
      );
      for (const group of data.value ?? []) {
        const matched = await groupIdIfSiteMatches(group);
        if (matched) {
          return matched;
        }
      }
    } catch {
      // Fall through to owned/member client groups.
    }
  }

  const groups = await listClientGroups(accessToken);
  for (const group of groups) {
    const matched = await groupIdIfSiteMatches(group);
    if (matched) {
      return matched;
    }
  }

  throw new Error(
    "Could not find the Microsoft 365 group for this Portal Link. Confirm the link is a CLIENT site created in the portal, then try again.",
  );
}

function isAlreadyGroupMemberError(message: string): boolean {
  return /already exist|added object references already exist|one or more added object references already exist/i.test(
    message,
  );
}

/**
 * Add directory users as members of the SharePoint site's M365 group
 * (resolved from Portal Link). Skips emails already in the group.
 */
export async function addSharePointSiteMembers(
  instance: IPublicClientApplication,
  account: AccountInfo,
  portalLink: string,
  memberEmails: string[],
): Promise<{ memberWarnings: string[] }> {
  const emails = [...new Set(memberEmails.map((e) => e.trim()).filter(Boolean))];
  if (emails.length === 0) {
    return { memberWarnings: [] };
  }

  const accessToken = await acquireSharePointToken(instance, account);
  const groupId = await resolveGroupIdFromPortalLink(accessToken, portalLink);
  const creatorEmail = String(
    account.username || account.name || "",
  ).toLowerCase();

  const memberWarnings: string[] = [];
  for (const email of emails) {
    if (email.toLowerCase() === creatorEmail) {
      continue;
    }
    try {
      await addGroupMember(accessToken, groupId, email);
    } catch (memberError) {
      const detail =
        memberError instanceof Error
          ? memberError.message
          : "Could not add member.";
      if (isAlreadyGroupMemberError(detail)) {
        continue;
      }
      memberWarnings.push(`${email}: ${detail}`);
    }
  }

  return { memberWarnings };
}

function parseSharePointSiteFromPortalLink(portalLink: string): {
  hostname: string;
  sitePath: string;
} | null {
  const trimmed = portalLink.trim();
  if (!trimmed) {
    return null;
  }
  let url: URL;
  try {
    url = new URL(/^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`);
  } catch {
    return null;
  }
  if (!/\.sharepoint\.com$/i.test(url.hostname)) {
    return null;
  }
  const match = url.pathname.match(/^(\/sites\/[^/]+)/i);
  if (!match?.[1]) {
    return null;
  }
  return { hostname: url.hostname, sitePath: match[1] };
}

type DriveChildrenPage = {
  value?: { name?: string; file?: unknown; folder?: unknown }[];
  "@odata.nextLink"?: string;
};

async function listDriveChildrenFileNames(
  accessToken: string,
  childrenUrl: string,
): Promise<string[]> {
  const names: string[] = [];
  let nextUrl: string | null = childrenUrl;
  let pages = 0;

  while (nextUrl && pages < 20) {
    pages += 1;
    const pageUrl = nextUrl;
    const data: DriveChildrenPage = await graphJson<DriveChildrenPage>(
      accessToken,
      pageUrl,
    );

    for (const item of data.value ?? []) {
      const name = String(item.name ?? "").trim();
      if (!name || item.folder) {
        continue;
      }
      names.push(name);
    }

    nextUrl = data["@odata.nextLink"]?.trim() || null;
  }

  return names;
}

/**
 * List file names in Shared Documents/{year}/{month} for the client's portal
 * site (same period folder BD1 email links to). Does not create folders.
 */
export async function listSharePointMonthFolderFiles(
  instance: IPublicClientApplication,
  account: AccountInfo,
  portalLink: string,
  forDate: Date = new Date(),
): Promise<{
  yearFolder: string;
  monthFolder: string;
  folderPath: string;
  fileNames: string[];
}> {
  const site = parseSharePointSiteFromPortalLink(portalLink);
  if (!site) {
    throw new Error(
      "Portal Link must be a SharePoint site URL (e.g. https://….sharepoint.com/sites/…).",
    );
  }

  const { yearFolder, monthFolder } = sharePointPeriodFolder(forDate);
  const folderPath = `Shared Documents/${yearFolder}/${monthFolder}`;
  const accessToken = await acquireSharePointToken(instance, account);

  const siteLookupPath = `${site.hostname}:${site.sitePath}`;
  const siteInfo = await graphJson<{ id?: string }>(
    accessToken,
    `${GRAPH}/sites/${siteLookupPath}?$select=id`,
  );
  const siteId = String(siteInfo.id ?? "").trim();
  if (!siteId) {
    throw new Error("Could not resolve the SharePoint site from Portal Link.");
  }

  const encodedFolder = `${yearFolder}/${monthFolder}`
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");

  try {
    const fileNames = await listDriveChildrenFileNames(
      accessToken,
      `${GRAPH}/sites/${siteId}/drive/root:/${encodedFolder}:/children?$select=name,file,folder&$top=200`,
    );
    return { yearFolder, monthFolder, folderPath, fileNames };
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to list folder files.";
    if (/itemNotFound|not found|404/i.test(message)) {
      throw new Error(
        `SharePoint folder not found: ${folderPath}. Create the year/month folders on the client site, then try again.`,
      );
    }
    throw error;
  }
}

export function sharePointFileNameMatches(
  fileNames: string[],
  needle: string,
): boolean {
  const target = needle.trim().toLowerCase();
  if (!target) {
    return false;
  }
  return fileNames.some((name) => name.toLowerCase().includes(target));
}
