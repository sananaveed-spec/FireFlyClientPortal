export type RequestBulletItem = {
  id: string;
  label: string;
  /** YYYY-MM when the bullet was added; same-month bullets appear on Action Items immediately. */
  createdMonth: string;
};

export type ActionItemStatus = "action" | "completed" | "late";

export type ActionWorkflowItem = {
  id: string;
  label: string;
  sourceRequestId: string;
  monthKey: string;
  status: ActionItemStatus;
  completedAt: string;
};

export type ClientRequestList = {
  clientId: string;
  scheduleId: string;
  requestItems: RequestBulletItem[];
  actionItems: ActionWorkflowItem[];
  lastPromotedMonth: string;
};

export type ClientRequestListsStore = {
  lists: ClientRequestList[];
};

function makeItemId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function currentMonthKey(date: Date = new Date()): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

export function createEmptyRequestListsStore(): ClientRequestListsStore {
  return { lists: [] };
}

function createActionFromRequest(
  item: RequestBulletItem,
  monthKey: string,
): ActionWorkflowItem {
  return {
    id: makeItemId("act"),
    label: item.label,
    sourceRequestId: item.id,
    monthKey,
    status: "action",
    completedAt: "",
  };
}

function hasActionForRequestMonth(
  actionItems: ActionWorkflowItem[],
  sourceRequestId: string,
  monthKey: string,
) {
  return actionItems.some(
    (item) =>
      item.sourceRequestId === sourceRequestId && item.monthKey === monthKey,
  );
}

export function normalizeRequestBulletItem(
  value: unknown,
): RequestBulletItem | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<RequestBulletItem> & { checked?: unknown };
  const label = String(record.label ?? "").trim();
  if (!label) {
    return null;
  }
  return {
    id: String(record.id ?? "").trim() || makeItemId("req"),
    label,
    createdMonth: String(record.createdMonth ?? "").trim(),
  };
}

export function normalizeActionWorkflowItem(
  value: unknown,
): ActionWorkflowItem | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<ActionWorkflowItem>;
  const label = String(record.label ?? "").trim();
  if (!label) {
    return null;
  }
  const status: ActionItemStatus =
    record.status === "completed"
      ? "completed"
      : record.status === "late"
        ? "late"
        : "action";
  return {
    id: String(record.id ?? "").trim() || makeItemId("act"),
    label,
    sourceRequestId: String(record.sourceRequestId ?? "").trim(),
    monthKey: String(record.monthKey ?? "").trim() || currentMonthKey(),
    status,
    completedAt:
      status === "completed" || status === "late"
        ? String(record.completedAt ?? "").trim() || new Date().toISOString()
        : "",
  };
}

export function normalizeClientRequestList(
  value: unknown,
): ClientRequestList | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Partial<ClientRequestList> & { items?: unknown };
  const clientId = String(record.clientId ?? "").trim();
  const scheduleId = String(record.scheduleId ?? "").trim();
  if (!clientId || !scheduleId) {
    return null;
  }

  const legacyItems = Array.isArray(record.items) ? record.items : [];
  const requestSource = Array.isArray(record.requestItems)
    ? record.requestItems
    : legacyItems;

  const requestItems = requestSource
    .map((item) => normalizeRequestBulletItem(item))
    .filter((item): item is RequestBulletItem => item !== null);

  const actionItems = Array.isArray(record.actionItems)
    ? record.actionItems
        .map((item) => normalizeActionWorkflowItem(item))
        .filter((item): item is ActionWorkflowItem => item !== null)
    : [];

  return {
    clientId,
    scheduleId,
    requestItems,
    actionItems,
    lastPromotedMonth: String(record.lastPromotedMonth ?? "").trim(),
  };
}

export function normalizeClientRequestListsStore(
  value: unknown,
): ClientRequestListsStore {
  if (!value || typeof value !== "object") {
    return createEmptyRequestListsStore();
  }
  const record = value as Partial<ClientRequestListsStore>;
  const lists = Array.isArray(record.lists)
    ? record.lists
        .map((item) => normalizeClientRequestList(item))
        .filter((item): item is ClientRequestList => item !== null)
    : [];
  return { lists };
}

export function getClientRequestList(
  store: ClientRequestListsStore,
  clientId: string,
  scheduleId: string,
): ClientRequestList {
  return (
    store.lists.find(
      (list) => list.clientId === clientId && list.scheduleId === scheduleId,
    ) ?? {
      clientId,
      scheduleId,
      requestItems: [],
      actionItems: [],
      lastPromotedMonth: "",
    }
  );
}

/**
 * Action Items / Completed are shared across every schedule slot (BD1, BD5, …)
 * for the same client. Dedupes by sourceRequestId + monthKey.
 */
export function getClientActionItems(
  store: ClientRequestListsStore,
  clientId: string,
): ActionWorkflowItem[] {
  const byKey = new Map<string, ActionWorkflowItem>();
  for (const list of store.lists) {
    if (list.clientId !== clientId) {
      continue;
    }
    for (const item of list.actionItems) {
      const key = `${item.sourceRequestId || item.id}::${item.monthKey}`;
      const existing = byKey.get(key);
      if (!existing) {
        byKey.set(key, item);
        continue;
      }
      // Prefer resolved (completed/late) when the same work item exists on
      // multiple schedules.
      if (existing.status === "action" && item.status !== "action") {
        byKey.set(key, item);
      }
    }
  }
  return Array.from(byKey.values());
}

export function findListContainingActionItem(
  store: ClientRequestListsStore,
  clientId: string,
  actionItemId: string,
): ClientRequestList | null {
  return (
    store.lists.find(
      (list) =>
        list.clientId === clientId &&
        list.actionItems.some((item) => item.id === actionItemId),
    ) ?? null
  );
}

/**
 * Request bullets created in the current month also appear on Action Items
 * immediately (without waiting for the next month).
 * Legacy bullets (no createdMonth) that already sit under this month's
 * lastPromotedMonth are treated as same-month creates.
 */
export function syncSameMonthCreationsToActionItems(
  list: ClientRequestList,
  now: Date = new Date(),
): ClientRequestList {
  const monthKey = currentMonthKey(now);
  let changed = false;

  const requestItems = list.requestItems.map((item) => {
    if (item.createdMonth) {
      return item;
    }
    // Backfill: items present while this month is already tracked were
    // created this month before createdMonth existed.
    if (list.lastPromotedMonth === monthKey) {
      changed = true;
      return { ...item, createdMonth: monthKey };
    }
    return item;
  });

  const missing = requestItems.filter(
    (item) =>
      item.createdMonth === monthKey &&
      !hasActionForRequestMonth(list.actionItems, item.id, monthKey),
  );

  if (!changed && missing.length === 0) {
    return list;
  }

  return {
    ...list,
    requestItems,
    actionItems: [
      ...list.actionItems,
      ...missing.map((item) => createActionFromRequest(item, monthKey)),
    ],
  };
}

/** Copy Request List bullets into Action Items when a new month starts. */
export function promoteRequestListForNewMonth(
  list: ClientRequestList,
  now: Date = new Date(),
): ClientRequestList {
  const monthKey = currentMonthKey(now);
  if (list.lastPromotedMonth === monthKey) {
    return syncSameMonthCreationsToActionItems(list, now);
  }
  // First time: start tracking this month; same-month creates still sync.
  if (!list.lastPromotedMonth) {
    return syncSameMonthCreationsToActionItems(
      { ...list, lastPromotedMonth: monthKey },
      now,
    );
  }

  const promoted: ActionWorkflowItem[] = list.requestItems
    .filter(
      (item) => !hasActionForRequestMonth(list.actionItems, item.id, monthKey),
    )
    .map((item) => createActionFromRequest(item, monthKey));

  return syncSameMonthCreationsToActionItems(
    {
      ...list,
      lastPromotedMonth: monthKey,
      actionItems: [...list.actionItems, ...promoted],
    },
    now,
  );
}

export function promoteAllListsForNewMonth(
  store: ClientRequestListsStore,
  now: Date = new Date(),
): { store: ClientRequestListsStore; changed: boolean } {
  let changed = false;
  const lists = store.lists.map((list) => {
    const next = promoteRequestListForNewMonth(list, now);
    if (next !== list) {
      changed = true;
    }
    return next;
  });
  return { store: { lists }, changed };
}

export function upsertClientRequestList(
  store: ClientRequestListsStore,
  nextList: ClientRequestList,
): ClientRequestListsStore {
  const normalized = normalizeClientRequestList(nextList);
  if (!normalized) {
    return store;
  }
  const synced = syncSameMonthCreationsToActionItems(normalized);
  const without = store.lists.filter(
    (list) =>
      !(
        list.clientId === synced.clientId &&
        list.scheduleId === synced.scheduleId
      ),
  );
  return { lists: [...without, synced] };
}

export function createRequestBulletItem(
  label: string,
  now: Date = new Date(),
): RequestBulletItem {
  return {
    id: makeItemId("req"),
    label: label.trim(),
    createdMonth: currentMonthKey(now),
  };
}

export function addRequestBulletToList(
  list: ClientRequestList,
  label: string,
  now: Date = new Date(),
): ClientRequestList {
  const bullet = createRequestBulletItem(label, now);
  const monthKey = currentMonthKey(now);
  return {
    ...list,
    requestItems: [...list.requestItems, bullet],
    actionItems: [
      ...list.actionItems,
      createActionFromRequest(bullet, monthKey),
    ],
  };
}

export function updateRequestBulletLabel(
  list: ClientRequestList,
  itemId: string,
  label: string,
): ClientRequestList {
  const nextLabel = label.trim();
  return {
    ...list,
    requestItems: list.requestItems.map((item) =>
      item.id === itemId ? { ...item, label: nextLabel } : item,
    ),
    actionItems: list.actionItems.map((item) =>
      item.sourceRequestId === itemId && item.status === "action"
        ? { ...item, label: nextLabel }
        : item,
    ),
  };
}

export function removeRequestBullet(
  list: ClientRequestList,
  itemId: string,
  now: Date = new Date(),
): ClientRequestList {
  const monthKey = currentMonthKey(now);
  return {
    ...list,
    requestItems: list.requestItems.filter((item) => item.id !== itemId),
    actionItems: list.actionItems.filter(
      (item) =>
        !(
          item.sourceRequestId === itemId &&
          item.monthKey === monthKey &&
          item.status === "action"
        ),
    ),
  };
}

export function completeActionItem(
  list: ClientRequestList,
  actionItemId: string,
  status: "completed" | "late" = "completed",
): ClientRequestList {
  return {
    ...list,
    actionItems: list.actionItems.map((item) =>
      item.id === actionItemId && item.status === "action"
        ? {
            ...item,
            status,
            completedAt: new Date().toISOString(),
          }
        : item,
    ),
  };
}

export function reopenActionItem(
  list: ClientRequestList,
  actionItemId: string,
): ClientRequestList {
  return {
    ...list,
    actionItems: list.actionItems.map((item) =>
      item.id === actionItemId &&
      (item.status === "completed" || item.status === "late")
        ? {
            ...item,
            status: "action",
            completedAt: "",
          }
        : item,
    ),
  };
}
