import { copyFile, mkdir, readFile, unlink, writeFile } from "fs/promises";
import path from "path";
import {
  migrateUniqueIds,
  normalizeClientProfile,
  type ClientProfileEntry,
} from "@/lib/client-profiles";
import {
  createHistoryId,
  normalizeHistoryEntry,
  type HistoryEntry,
  type HistoryInput,
} from "@/lib/history";
import {
  createEmptyEmailSignature,
  normalizeEmailSignaturesStore,
  type EmailSignature,
  type EmailSignaturesStore,
} from "@/lib/email-signatures";
import {
  createDefaultEmailTemplates,
  normalizeEmailTemplates,
  type EmailTemplatesStore,
} from "@/lib/email-templates";
import { normalizeEmail } from "@/lib/portal-users";
import {
  createDefaultHolidays,
  normalizeHolidays,
  type HolidaysStore,
} from "@/lib/holidays";
import {
  createDefaultLookupData,
  normalizeLookupData,
  type LookupDataStore,
} from "@/lib/lookup-data";
import {
  normalizeResourceClient,
  type ResourceClientEntry,
} from "@/lib/resource-clients";

/** Always the project `data` folder — survives app restarts and reopen. */
export function getDataDir() {
  return process.env.DATA_DIR?.trim() || path.join(process.cwd(), "data");
}

async function ensureDataDir() {
  await mkdir(getDataDir(), { recursive: true });
}

function clientProfilesPath() {
  return path.join(getDataDir(), "client-profiles.json");
}

function resourceClientsPath() {
  return path.join(getDataDir(), "resource-clients.json");
}

function lookupDataPath() {
  return path.join(getDataDir(), "lookup-data.json");
}

function holidaysPath() {
  return path.join(getDataDir(), "holidays.json");
}

function emailTemplatesPath() {
  return path.join(getDataDir(), "email-templates.json");
}

function emailSignaturesPath() {
  return path.join(getDataDir(), "email-signatures.json");
}

function historyPath() {
  return path.join(getDataDir(), "history.json");
}

function portalUsersPath() {
  return path.join(getDataDir(), "portal-users.json");
}

function clientRequestListsPath() {
  return path.join(getDataDir(), "client-request-lists.json");
}

async function writeJsonFile(filePath: string, value: unknown) {
  const tempPath = `${filePath}.${process.pid}.tmp`;
  const payload = `${JSON.stringify(value, null, 2)}\n`;
  await writeFile(tempPath, payload, "utf8");
  await copyFile(tempPath, filePath);
  await unlink(tempPath).catch(() => undefined);
}

function parseFirstJsonValue(text: string): unknown {
  const source = text.replace(/^\uFEFF/, "").trimStart();
  let depth = 0;
  let inString = false;
  let escape = false;

  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (inString) {
      if (escape) {
        escape = false;
      } else if (char === "\\") {
        escape = true;
      } else if (char === '"') {
        inString = false;
      }
      continue;
    }
    if (char === '"') {
      inString = true;
      continue;
    }
    if (char === "{" || char === "[") {
      depth += 1;
    } else if (char === "}" || char === "]") {
      depth -= 1;
      if (depth === 0) {
        return JSON.parse(source.slice(0, index + 1));
      }
    }
  }

  return JSON.parse(source);
}

function parseJsonFile(raw: string): unknown {
  // Strip UTF-8 BOM if present (can appear after some Windows editors).
  try {
    return JSON.parse(raw.replace(/^\uFEFF/, ""));
  } catch {
    // Recover when a file has trailing garbage after valid JSON.
    return parseFirstJsonValue(raw);
  }
}

function isMissingFile(error: unknown) {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    (error as { code?: string }).code === "ENOENT"
  );
}

export async function readClientProfiles(): Promise<ClientProfileEntry[]> {
  await ensureDataDir();
  const filePath = clientProfilesPath();

  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = parseJsonFile(raw);
    if (!Array.isArray(parsed)) {
      throw new Error("Invalid client-profiles.json format.");
    }

    const profiles = parsed
      .filter(
        (item): item is Partial<ClientProfileEntry> & { id: string } =>
          typeof item === "object" &&
          item !== null &&
          typeof (item as { id?: unknown }).id === "string",
      )
      .map(normalizeClientProfile);

    const migrated = migrateUniqueIds(profiles);
    if (migrated.changed) {
      await writeJsonFile(filePath, migrated.profiles);
    }
    return migrated.profiles;
  } catch (error) {
    if (isMissingFile(error)) {
      await writeClientProfiles([]);
      return [];
    }
    throw error;
  }
}

export async function writeClientProfiles(
  profiles: ClientProfileEntry[],
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(clientProfilesPath(), profiles);
}

export async function readResourceClients(): Promise<ResourceClientEntry[]> {
  await ensureDataDir();
  const filePath = resourceClientsPath();

  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = parseJsonFile(raw);
    if (!Array.isArray(parsed)) {
      throw new Error("Invalid resource-clients.json format.");
    }

    return parsed
      .filter(
        (item): item is Partial<ResourceClientEntry> & { id: string } =>
          typeof item === "object" &&
          item !== null &&
          typeof (item as { id?: unknown }).id === "string",
      )
      .map(normalizeResourceClient);
  } catch (error) {
    if (isMissingFile(error)) {
      await writeResourceClients([]);
      return [];
    }
    throw error;
  }
}

export async function writeResourceClients(
  clients: ResourceClientEntry[],
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(resourceClientsPath(), clients);
}

export async function readLookupData(): Promise<LookupDataStore> {
  await ensureDataDir();
  const filePath = lookupDataPath();

  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = parseJsonFile(raw);
    return normalizeLookupData(parsed);
  } catch (error) {
    if (isMissingFile(error)) {
      const defaults = createDefaultLookupData();
      await writeLookupData(defaults);
      return defaults;
    }
    throw error;
  }
}

export async function writeLookupData(data: LookupDataStore): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(lookupDataPath(), normalizeLookupData(data));
}

export async function readHolidays(): Promise<HolidaysStore> {
  await ensureDataDir();
  const filePath = holidaysPath();

  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = parseJsonFile(raw);
    return normalizeHolidays(parsed);
  } catch (error) {
    if (isMissingFile(error)) {
      const defaults = createDefaultHolidays();
      await writeHolidays(defaults);
      return defaults;
    }
    throw error;
  }
}

export async function writeHolidays(data: HolidaysStore): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(holidaysPath(), normalizeHolidays(data));
}

export async function readEmailTemplates(): Promise<EmailTemplatesStore> {
  await ensureDataDir();
  const filePath = emailTemplatesPath();

  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = parseJsonFile(raw);
    return normalizeEmailTemplates(parsed);
  } catch (error) {
    if (isMissingFile(error)) {
      const defaults = createDefaultEmailTemplates();
      await writeEmailTemplates(defaults);
      return defaults;
    }
    throw error;
  }
}

export async function writeEmailTemplates(
  data: EmailTemplatesStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(emailTemplatesPath(), normalizeEmailTemplates(data));
}

export async function readEmailSignatures(): Promise<EmailSignaturesStore> {
  await ensureDataDir();
  const filePath = emailSignaturesPath();

  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = parseJsonFile(raw);
    return normalizeEmailSignaturesStore(parsed);
  } catch (error) {
    if (isMissingFile(error)) {
      const empty: EmailSignaturesStore = { signatures: {} };
      await writeEmailSignatures(empty);
      return empty;
    }
    throw error;
  }
}

export async function writeEmailSignatures(
  data: EmailSignaturesStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(
    emailSignaturesPath(),
    normalizeEmailSignaturesStore(data),
  );
}

export async function readEmailSignatureForUser(
  email: string,
): Promise<EmailSignature> {
  const store = await readEmailSignatures();
  const key = normalizeEmail(email);
  return store.signatures[key] ?? createEmptyEmailSignature();
}

export async function writeEmailSignatureForUser(
  email: string,
  signature: EmailSignature,
): Promise<EmailSignature> {
  const key = normalizeEmail(email);
  const store = await readEmailSignatures();
  const nextSignature = {
    ...signature,
    updatedAt: new Date().toISOString(),
  };
  store.signatures[key] = nextSignature;
  await writeEmailSignatures(store);
  return nextSignature;
}

export async function readHistory(): Promise<HistoryEntry[]> {
  await ensureDataDir();
  const filePath = historyPath();

  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = parseJsonFile(raw);
    if (!Array.isArray(parsed)) {
      throw new Error("Invalid history.json format.");
    }

    return parsed
      .map((item) =>
        typeof item === "object" && item !== null
          ? normalizeHistoryEntry(item as Partial<HistoryEntry>)
          : null,
      )
      .filter((item): item is HistoryEntry => item !== null)
      .sort(
        (a, b) =>
          new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
      );
  } catch (error) {
    if (isMissingFile(error)) {
      await writeHistory([]);
      return [];
    }
    throw error;
  }
}

export async function writeHistory(entries: HistoryEntry[]): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(historyPath(), entries);
}

export async function appendHistory(
  input: HistoryInput,
): Promise<HistoryEntry> {
  const entries = await readHistory();
  const entry = normalizeHistoryEntry({
    ...input,
    id: createHistoryId(),
    createdAt: input.createdAt ?? new Date().toISOString(),
  });

  if (!entry) {
    throw new Error("Invalid history entry.");
  }

  const next = [entry, ...entries];
  await writeHistory(next);
  return entry;
}

export async function readPortalUsersFile(): Promise<unknown> {
  await ensureDataDir();
  const filePath = portalUsersPath();

  try {
    const raw = await readFile(filePath, "utf8");
    return parseJsonFile(raw);
  } catch (error) {
    if (isMissingFile(error)) {
      return [];
    }
    throw error;
  }
}

export async function writePortalUsersFile(users: unknown): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(portalUsersPath(), users);
}

export async function readClientRequestListsFile(): Promise<unknown> {
  await ensureDataDir();
  const filePath = clientRequestListsPath();

  try {
    const raw = await readFile(filePath, "utf8");
    return parseJsonFile(raw);
  } catch (error) {
    if (isMissingFile(error)) {
      return { lists: [] };
    }
    throw error;
  }
}

export async function writeClientRequestListsFile(
  value: unknown,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(clientRequestListsPath(), value);
}
