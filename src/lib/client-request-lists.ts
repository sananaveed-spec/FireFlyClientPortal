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

/** Mid-month Date for a YYYY-MM key (stable for month-scoped action items). */
export function parseMonthKeyToClock(month: string): Date | null {
  if (!/^\d{4}-\d{2}$/.test(month.trim())) {
    return null;
  }
  const [year, mon] = month.trim().split("-").map(Number);
  if (!year || !mon || mon < 1 || mon > 12) {
    return null;
  }
  return new Date(year, mon - 1, 15, 12, 0, 0, 0);
}

/** Request bullets that belong to a calendar month (BD1 / BD8 lists). */
export function requestItemsForMonth(
  items: RequestBulletItem[],
  monthKey: string,
): RequestBulletItem[] {
  return items.filter((item) => item.createdMonth === monthKey);
}

/**
 * Pin legacy bullets (no createdMonth) to the prior tracked month so a demo
 * or rollover cannot re-tag them as belonging to the new month's BD1/BD8 form.
 */
function pinLegacyRequestMonths(list: ClientRequestList): ClientRequestList {
  if (!list.lastPromotedMonth) {
    return list;
  }
  let changed = false;
  const requestItems = list.requestItems.map((item) => {
    if (item.createdMonth) {
      return item;
    }
    changed = true;
    return { ...item, createdMonth: list.lastPromotedMonth };
  });
  return changed ? { ...list, requestItems } : list;
}

/**
 * Month of the user's latest request-list edits (max YYYY-MM that has bullets).
 * Month switches always project from this list — not an older month.
 */
function latestEditedRequestMonth(items: RequestBulletItem[]): string | null {
  const months = [
    ...new Set(
      items
        .map((item) => item.createdMonth)
        .filter((month): month is string => Boolean(month)),
    ),
  ].sort();
  return months.length > 0 ? (months[months.length - 1] ?? null) : null;
}

function requestLabelFingerprint(items: RequestBulletItem[]): string {
  return items
    .map((item) => item.label.trim().toLowerCase())
    .filter(Boolean)
    .join("\n");
}

/**
 * Ensure `monthKey` shows the last list the user edited (any prior month).
 * Example: Oct=2, Nov=3 → Dec gets Nov's 3 (not Oct's 2).
 */
function copyRequestBulletsToMonth(
  list: ClientRequestList,
  monthKey: string,
): ClientRequestList {
  const sourceMonth = latestEditedRequestMonth(list.requestItems);
  if (!sourceMonth || sourceMonth === monthKey) {
    return list;
  }
  const sourceItems = list.requestItems.filter(
    (item) => item.createdMonth === sourceMonth,
  );
  if (sourceItems.length === 0) {
    return list;
  }
  const targetItems = list.requestItems.filter(
    (item) => item.createdMonth === monthKey,
  );
  if (
    targetItems.length > 0 &&
    requestLabelFingerprint(targetItems) ===
      requestLabelFingerprint(sourceItems)
  ) {
    return list;
  }

  const targetIds = new Set(targetItems.map((item) => item.id));
  const clones = sourceItems.map((item) => ({
    id: makeItemId("req"),
    label: item.label,
    createdMonth: monthKey,
  }));

  return {
    ...list,
    requestItems: [
      ...list.requestItems.filter((item) => item.createdMonth !== monthKey),
      ...clones,
    ],
    // Drop Open rows that belonged to the replaced target-month bullets.
    actionItems: list.actionItems.filter(
      (action) =>
        !(
          action.monthKey === monthKey &&
          action.status === "action" &&
          targetIds.has(action.sourceRequestId)
        ),
    ),
  };
}

/**
 * Drop Open rows for `monthKey` that still point at another month's request
 * bullet (legacy promote), and collapse duplicate Open labels for the month.
 */
function cleanupMonthOpenActions(
  list: ClientRequestList,
  monthKey: string,
): ClientRequestList {
  const requestById = new Map(
    list.requestItems.map((item) => [item.id, item] as const),
  );
  const kept: ActionWorkflowItem[] = [];
  const openLabels = new Set<string>();

  for (const action of list.actionItems) {
    if (action.monthKey !== monthKey) {
      kept.push(action);
      continue;
    }
    if (action.status !== "action") {
      kept.push(action);
      continue;
    }
    const source = requestById.get(action.sourceRequestId);
    if (source && source.createdMonth && source.createdMonth !== monthKey) {
      // Stale Open row from prior-month source — Nov copy owns this work now.
      continue;
    }
    const labelKey = action.label.trim().toLowerCase();
    if (labelKey && openLabels.has(labelKey)) {
      continue;
    }
    if (labelKey) {
      openLabels.add(labelKey);
    }
    kept.push(action);
  }

  if (kept.length === list.actionItems.length) {
    return list;
  }
  return { ...list, actionItems: kept };
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
      const labelKey = item.label.trim().toLowerCase() || item.id;
      const key = `${labelKey}::${item.monthKey}`;
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
 * Request bullets created in the active month also appear on Action Items
 * immediately. Prior-month bullets are never re-tagged or re-opened.
 */
export function syncSameMonthCreationsToActionItems(
  list: ClientRequestList,
  now: Date = new Date(),
): ClientRequestList {
  const monthKey = currentMonthKey(now);
  const missing = list.requestItems.filter(
    (item) =>
      item.createdMonth === monthKey &&
      !hasActionForRequestMonth(list.actionItems, item.id, monthKey),
  );

  if (missing.length === 0) {
    return list;
  }

  return {
    ...list,
    actionItems: [
      ...list.actionItems,
      ...missing.map((item) => createActionFromRequest(item, monthKey)),
    ],
  };
}

/**
 * Advance the tracked month: project the user's latest BD1/BD8 list edits
 * onto this month, then open matching Action rows (BD5 / BD9).
 */
export function promoteRequestListForNewMonth(
  list: ClientRequestList,
  now: Date = new Date(),
): ClientRequestList {
  const monthKey = currentMonthKey(now);
  let next = pinLegacyRequestMonths(list);
  next = copyRequestBulletsToMonth(next, monthKey);

  if (next.lastPromotedMonth === monthKey) {
    return cleanupMonthOpenActions(
      syncSameMonthCreationsToActionItems(next, now),
      monthKey,
    );
  }

  // First time: start tracking this month; same-month creates still sync.
  if (!next.lastPromotedMonth) {
    const seeded = {
      ...next,
      requestItems: next.requestItems.map((item) =>
        item.createdMonth ? item : { ...item, createdMonth: monthKey },
      ),
      lastPromotedMonth: monthKey,
    };
    return cleanupMonthOpenActions(
      syncSameMonthCreationsToActionItems(seeded, now),
      monthKey,
    );
  }

  // Never move lastPromotedMonth backwards (demo month can jump around).
  const nextLast =
    next.lastPromotedMonth > monthKey ? next.lastPromotedMonth : monthKey;

  return cleanupMonthOpenActions(
    syncSameMonthCreationsToActionItems(
      {
        ...next,
        lastPromotedMonth: nextLast,
      },
      now,
    ),
    monthKey,
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
  now: Date = new Date(),
): ClientRequestListsStore {
  const normalized = normalizeClientRequestList(nextList);
  if (!normalized) {
    return store;
  }
  const synced = syncSameMonthCreationsToActionItems(normalized, now);
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

/** Set Completed/Late even if the item was already resolved (e.g. SharePoint re-check). */
export function setActionItemStatus(
  list: ClientRequestList,
  actionItemId: string,
  status: "completed" | "late",
): ClientRequestList {
  return {
    ...list,
    actionItems: list.actionItems.map((item) =>
      item.id === actionItemId
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
