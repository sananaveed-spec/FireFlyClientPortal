export type AdjustableSettingRow = {
  id: string;
  templateId: string;
  businessDayId: string;
};

export type ClientAdjustableSettings = {
  clientId: string;
  rows: AdjustableSettingRow[];
};

export type FlexibleAdjustableSettingsStore = {
  clients: ClientAdjustableSettings[];
};

function makeRowId() {
  return `adj-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function createEmptyAdjustableSettingsStore(): FlexibleAdjustableSettingsStore {
  return { clients: [] };
}

export function normalizeAdjustableSettingRow(
  value: unknown,
): AdjustableSettingRow | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<AdjustableSettingRow>;
  const templateId = String(record.templateId ?? "").trim();
  const businessDayId = String(record.businessDayId ?? "").trim();
  if (!templateId || !businessDayId) {
    return null;
  }
  const id = String(record.id ?? "").trim() || makeRowId();
  return { id, templateId, businessDayId };
}

export function normalizeClientAdjustableSettings(
  value: unknown,
): ClientAdjustableSettings | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<ClientAdjustableSettings>;
  const clientId = String(record.clientId ?? "").trim();
  if (!clientId) {
    return null;
  }
  const rows = Array.isArray(record.rows)
    ? record.rows
        .map(normalizeAdjustableSettingRow)
        .filter((row): row is AdjustableSettingRow => Boolean(row))
    : [];
  return { clientId, rows };
}

export function normalizeFlexibleAdjustableSettingsStore(
  value: unknown,
): FlexibleAdjustableSettingsStore {
  if (!value || typeof value !== "object") {
    return createEmptyAdjustableSettingsStore();
  }
  const record = value as Partial<FlexibleAdjustableSettingsStore>;
  const clients = Array.isArray(record.clients)
    ? record.clients
        .map(normalizeClientAdjustableSettings)
        .filter((entry): entry is ClientAdjustableSettings => Boolean(entry))
    : [];
  return { clients };
}

export function getClientAdjustableRows(
  store: FlexibleAdjustableSettingsStore,
  clientId: string,
): AdjustableSettingRow[] {
  const id = clientId.trim();
  if (!id) {
    return [];
  }
  return store.clients.find((entry) => entry.clientId === id)?.rows ?? [];
}

export function upsertClientAdjustableRows(
  store: FlexibleAdjustableSettingsStore,
  clientId: string,
  rows: AdjustableSettingRow[],
): FlexibleAdjustableSettingsStore {
  const id = clientId.trim();
  if (!id) {
    return store;
  }
  const normalizedRows = rows
    .map(normalizeAdjustableSettingRow)
    .filter((row): row is AdjustableSettingRow => Boolean(row));
  const without = store.clients.filter((entry) => entry.clientId !== id);
  return {
    clients: [...without, { clientId: id, rows: normalizedRows }],
  };
}

export function createAdjustableSettingRow(input: {
  templateId: string;
  businessDayId: string;
}): AdjustableSettingRow | null {
  return normalizeAdjustableSettingRow({
    id: makeRowId(),
    templateId: input.templateId,
    businessDayId: input.businessDayId,
  });
}
