export type ClientScheduleSettings = {
  clientId: string;
  /** Business-day schedule ids with a filled star (Active). */
  activeScheduleIds: string[];
};

export type ClientScheduleSettingsStore = {
  settings: ClientScheduleSettings[];
};

export function createEmptyScheduleSettingsStore(): ClientScheduleSettingsStore {
  return { settings: [] };
}

export function normalizeClientScheduleSettings(
  value: unknown,
): ClientScheduleSettings | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<ClientScheduleSettings>;
  const clientId = String(record.clientId ?? "").trim();
  if (!clientId) {
    return null;
  }
  const activeScheduleIds = Array.isArray(record.activeScheduleIds)
    ? [
        ...new Set(
          record.activeScheduleIds
            .map((id) => {
              const normalized = String(id ?? "").trim();
              // Renamed schedule slot: Questions + 3 BD → BD9
              if (normalized === "questions-plus-3-bd") {
                return "bd9";
              }
              // Merged Last N BD (Team) into Last N BD + Client/Team toggle
              const teamMatch = /^last-(\d+)-bd-team$/i.exec(normalized);
              if (teamMatch) {
                return `last-${teamMatch[1]}-bd`;
              }
              return normalized;
            })
            .filter(Boolean),
        ),
      ]
    : [];
  return { clientId, activeScheduleIds };
}

export function normalizeClientScheduleSettingsStore(
  value: unknown,
): ClientScheduleSettingsStore {
  if (!value || typeof value !== "object") {
    return createEmptyScheduleSettingsStore();
  }
  const record = value as Partial<ClientScheduleSettingsStore>;
  const settings = Array.isArray(record.settings)
    ? record.settings
        .map((item) => normalizeClientScheduleSettings(item))
        .filter((item): item is ClientScheduleSettings => Boolean(item))
    : [];
  return { settings };
}

export function getClientScheduleSettings(
  store: ClientScheduleSettingsStore,
  clientId: string,
): ClientScheduleSettings | null {
  const id = clientId.trim();
  if (!id) {
    return null;
  }
  return store.settings.find((entry) => entry.clientId === id) ?? null;
}

/**
 * Active schedule ids for a client. When no saved settings exist yet,
 * every known schedule id is treated as active (filled star).
 */
export function getActiveScheduleIds(
  store: ClientScheduleSettingsStore,
  clientId: string,
  allScheduleIds: readonly string[],
): string[] {
  const existing = getClientScheduleSettings(store, clientId);
  if (!existing) {
    return [...allScheduleIds];
  }
  const allowed = new Set(allScheduleIds);
  return existing.activeScheduleIds.filter((id) => allowed.has(id));
}

export function isScheduleActiveForClient(
  store: ClientScheduleSettingsStore,
  clientId: string,
  scheduleId: string,
  allScheduleIds: readonly string[],
): boolean {
  return getActiveScheduleIds(store, clientId, allScheduleIds).includes(
    scheduleId,
  );
}

export function upsertClientScheduleSettings(
  store: ClientScheduleSettingsStore,
  next: ClientScheduleSettings,
): ClientScheduleSettingsStore {
  const normalized = normalizeClientScheduleSettings(next);
  if (!normalized) {
    return store;
  }
  const without = store.settings.filter(
    (entry) => entry.clientId !== normalized.clientId,
  );
  return { settings: [...without, normalized] };
}

export function toggleClientScheduleActive(
  store: ClientScheduleSettingsStore,
  clientId: string,
  scheduleId: string,
  allScheduleIds: readonly string[],
): ClientScheduleSettingsStore {
  const id = clientId.trim();
  const schedule = scheduleId.trim();
  if (!id || !schedule || !allScheduleIds.includes(schedule)) {
    return store;
  }

  const current = new Set(getActiveScheduleIds(store, id, allScheduleIds));
  if (current.has(schedule)) {
    current.delete(schedule);
  } else {
    current.add(schedule);
  }

  return upsertClientScheduleSettings(store, {
    clientId: id,
    activeScheduleIds: allScheduleIds.filter((item) => current.has(item)),
  });
}
