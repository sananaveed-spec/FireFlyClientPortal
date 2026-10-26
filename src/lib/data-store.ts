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
  createDefaultRepeatedTasks,
  normalizeRepeatedTasks,
  type RepeatedTasksStore,
} from "@/lib/repeated-tasks";
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

function repeatedTasksPath() {
  return path.join(getDataDir(), "repeated-tasks.json");
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

function clientScheduleSettingsPath() {
  return path.join(getDataDir(), "client-schedule-settings.json");
}

function bd8EmailAttachmentsPath() {
  return path.join(getDataDir(), "bd8-email-attachments.json");
}

function bd15EmailAttachmentsPath() {
  return path.join(getDataDir(), "bd15-email-attachments.json");
}

function bd15EmailStatusPath() {
  return path.join(getDataDir(), "bd15-email-status.json");
}

function pakistanHolidayEmailsPath() {
  return path.join(getDataDir(), "pakistan-holiday-emails.json");
}

function ignitionEndingSoonEmailsPath() {
  return path.join(getDataDir(), "ignition-ending-soon-emails.json");
}

async function writeJsonFile(filePath: string, value: unknown) {
  const serialized = JSON.stringify(value, null, 2);
  if (typeof serialized !== "string" || !serialized.trim()) {
    throw new Error(`Refusing to write empty JSON to ${path.basename(filePath)}.`);
  }
  const payload = `${serialized}\n`;
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await writeFile(tempPath, payload, "utf8");
  // Confirm temp is valid before replacing the live file (avoids 0-byte wipes).
  const verify = await readFile(tempPath, "utf8");
  JSON.parse(verify);
  await copyFile(tempPath, filePath);
  await unlink(tempPath).catch(() => undefined);
}

/** Read JSON; treat empty/corrupt as missing so callers can reseeds defaults. */
async function readJsonFileOrNull(filePath: string): Promise<unknown | null> {
  try {
    const raw = await readFile(filePath, "utf8");
    if (!raw.trim()) {
      return null;
    }
    try {
      return parseJsonFile(raw);
    } catch {
      return null;
    }
  } catch (error) {
    if (isMissingFile(error)) {
      return null;
    }
    throw error;
  }
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

export async function readRepeatedTasks(): Promise<RepeatedTasksStore> {
  await ensureDataDir();
  const filePath = repeatedTasksPath();
  const parsed = await readJsonFileOrNull(filePath);
  if (parsed == null) {
    const defaults = createDefaultRepeatedTasks();
    await writeRepeatedTasks(defaults);
    return defaults;
  }
  return normalizeRepeatedTasks(parsed);
}

export async function writeRepeatedTasks(
  data: RepeatedTasksStore,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(repeatedTasksPath(), normalizeRepeatedTasks(data));
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
  const parsed = await readJsonFileOrNull(filePath);
  if (parsed == null) {
    const empty = { lists: [] };
    await writeJsonFile(filePath, empty);
    return empty;
  }
  return parsed;
}

export async function writeClientRequestListsFile(
  value: unknown,
): Promise<void> {
  await ensureDataDir();
  // Never persist a non-object / undefined payload (can truncate the file).
  if (!value || typeof value !== "object") {
    throw new Error("Invalid client request lists payload.");
  }
  await writeJsonFile(clientRequestListsPath(), value);
}

export async function readClientScheduleSettingsFile(): Promise<unknown> {
  await ensureDataDir();
  const filePath = clientScheduleSettingsPath();

  try {
    const raw = await readFile(filePath, "utf8");
    return parseJsonFile(raw);
  } catch (error) {
    if (isMissingFile(error)) {
      return { settings: [] };
    }
    throw error;
  }
}

export async function writeClientScheduleSettingsFile(
  value: unknown,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(clientScheduleSettingsPath(), value);
}

export async function readBd8EmailAttachmentsFile(): Promise<unknown> {
  await ensureDataDir();
  const filePath = bd8EmailAttachmentsPath();

  try {
    const raw = await readFile(filePath, "utf8");
    return parseJsonFile(raw);
  } catch (error) {
    if (isMissingFile(error)) {
      return { attachments: [] };
    }
    throw error;
  }
}

export async function writeBd8EmailAttachmentsFile(
  value: unknown,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(bd8EmailAttachmentsPath(), value);
}

export async function readBd15EmailAttachmentsFile(): Promise<unknown> {
  await ensureDataDir();
  const filePath = bd15EmailAttachmentsPath();

  try {
    const raw = await readFile(filePath, "utf8");
    return parseJsonFile(raw);
  } catch (error) {
    if (isMissingFile(error)) {
      return { attachments: [] };
    }
    throw error;
  }
}

export async function writeBd15EmailAttachmentsFile(
  value: unknown,
): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(bd15EmailAttachmentsPath(), value);
}

export async function readBd15EmailStatusFile(): Promise<unknown> {
  await ensureDataDir();
  const filePath = bd15EmailStatusPath();

  try {
    const raw = await readFile(filePath, "utf8");
    return parseJsonFile(raw);
  } catch (error) {
    if (isMissingFile(error)) {
      return { statuses: [] };
    }
    throw error;
  }
}

export async function writeBd15EmailStatusFile(value: unknown): Promise<void> {
  await ensureDataDir();
  await writeJsonFile(bd15EmailStatusPath(), value);
}

export async function readPakistanHolidayEmailsFile(): Promise<unknown> {
  await ensureDataDir();
  const filePath = pakistanHolidayEmailsPath();
  const parsed = await readJsonFileOrNull(filePath);
  if (parsed == null) {
    const empty = { entries: [] };
    await writeJsonFile(filePath, empty);
    return empty;
  }
  return parsed;
}

export async function writePakistanHolidayEmailsFile(
  value: unknown,
): Promise<void> {
  await ensureDataDir();
  if (!value || typeof value !== "object") {
    throw new Error("Invalid Pakistan holiday emails payload.");
  }
  await writeJsonFile(pakistanHolidayEmailsPath(), value);
}

export async function readIgnitionEndingSoonEmailsFile(): Promise<unknown> {
  await ensureDataDir();
  const filePath = ignitionEndingSoonEmailsPath();
  const parsed = await readJsonFileOrNull(filePath);
  if (parsed == null) {
    const empty = { dryRun: false, entries: [] };
    await writeJsonFile(filePath, empty);
    return empty;
  }
  return parsed;
}

export async function writeIgnitionEndingSoonEmailsFile(
  value: unknown,
): Promise<void> {
  await ensureDataDir();
  if (!value || typeof value !== "object") {
    throw new Error("Invalid Ignition ending-soon emails payload.");
  }
  await writeJsonFile(ignitionEndingSoonEmailsPath(), value);
}
