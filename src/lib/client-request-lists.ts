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

/** BD5 schedule id paired with a BD1 request-type list. */
export function bd5ScheduleIdForBd1List(bd1ScheduleId: string): string {
  return bd1ScheduleId === "bd1-other-statements"
    ? "bd5-other-statements"
    : "bd5";
}

export function bd1ScheduleIdForRequestType(
  otherStatements: boolean,
): "bd1" | "bd1-other-statements" {
  return otherStatements ? "bd1-other-statements" : "bd1";
}

/**
 * After BD1 email send: freeze that month’s request labels onto BD5
 * (non-editable snapshot). First snapshot for the month wins — BD1 is
 * once-per-month. Late/Complete on BD5 never writes back to BD1.
 *
 * `preserveStatusesFromBd1`: for backfill when Late was already marked on
 * the shared BD1 rows before snapshots existed.
 */
export function snapshotBd1ToBd5ForMonth(
  store: ClientRequestListsStore,
  clientId: string,
  monthKey: string,
  options: {
    otherStatements: boolean;
    preserveStatusesFromBd1?: boolean;
  },
): { store: ClientRequestListsStore; created: boolean } {
  const id = clientId.trim();
  const month = monthKey.trim();
  if (!id || !/^\d{4}-\d{2}$/.test(month)) {
    return { store, created: false };
  }

  const sourceId = bd1ScheduleIdForRequestType(options.otherStatements);
  const targetId = bd5ScheduleIdForBd1List(sourceId);
  const source = getClientRequestList(store, id, sourceId);
  const target = getClientRequestList(store, id, targetId);

  const existingBd5ForMonth = target.actionItems.filter(
    (item) => item.monthKey === month,
  );
  if (existingBd5ForMonth.length > 0) {
    return { store, created: false };
  }

  const periodRequests = requestItemsForMonth(source.requestItems, month);
  const labelSources =
    periodRequests.length > 0
      ? periodRequests
      : source.requestItems.filter((item) => item.label.trim());
  // Prefer unique labels in list order.
  const labels: string[] = [];
  const seen = new Set<string>();
  for (const item of labelSources) {
    const label = item.label.trim();
    const key = label.toLowerCase();
    if (!label || seen.has(key)) {
      continue;
    }
    seen.add(key);
    labels.push(label);
  }
  // Fall back to BD1 action labels for the month if request bullets missing.
  if (labels.length === 0) {
    for (const item of source.actionItems) {
      if (item.monthKey !== month) {
        continue;
      }
      const label = item.label.trim();
      const key = label.toLowerCase();
      if (!label || seen.has(key)) {
        continue;
      }
      seen.add(key);
      labels.push(label);
    }
  }
  if (labels.length === 0) {
    return { store, created: false };
  }

  const bd1ByLabel = new Map<string, ActionWorkflowItem>();
  if (options.preserveStatusesFromBd1) {
    for (const item of source.actionItems) {
      if (item.monthKey !== month) {
        continue;
      }
      const key = item.label.trim().toLowerCase();
      if (!key) {
        continue;
      }
      const existing = bd1ByLabel.get(key);
      if (
        !existing ||
        (existing.status === "action" && item.status !== "action")
      ) {
        bd1ByLabel.set(key, item);
      }
    }
  }

  const newRequests: RequestBulletItem[] = labels.map((label) => ({
    id: makeItemId("req"),
    label,
    createdMonth: month,
  }));
  const newActions: ActionWorkflowItem[] = newRequests.map((req) => {
    const fromBd1 = bd1ByLabel.get(req.label.toLowerCase());
    const status =
      options.preserveStatusesFromBd1 && fromBd1
        ? fromBd1.status
        : ("action" as const);
    return {
      id: makeItemId("act"),
      label: req.label,
      sourceRequestId: `bd5-snap:${req.id}`,
      monthKey: month,
      status,
      completedAt:
        status === "action"
          ? ""
          : fromBd1?.completedAt || new Date().toISOString(),
    };
  });

  const nextTarget: ClientRequestList = {
    clientId: id,
    scheduleId: targetId,
    requestItems: [
      ...target.requestItems.filter((item) => item.createdMonth !== month),
      ...newRequests,
    ],
    actionItems: [
      ...target.actionItems.filter((item) => item.monthKey !== month),
      ...newActions,
    ],
    lastPromotedMonth: target.lastPromotedMonth || month,
  };

  return {
    store: upsertClientRequestList(store, nextTarget),
    created: true,
  };
}

/** BD5 Late only (what BD7 / cron use) — never BD1. */
export function getBd5LateActionItemsForMonth(
  store: ClientRequestListsStore,
  clientId: string,
  monthKey: string,
): ActionWorkflowItem[] {
  const id = clientId.trim();
  const month = monthKey.trim();
  if (!id || !month) {
    return [];
  }
  const lateByKey = new Map<string, ActionWorkflowItem>();
  for (const scheduleId of ["bd5", "bd5-other-statements"]) {
    const list = getClientRequestList(store, id, scheduleId);
    for (const item of list.actionItems) {
      if (item.monthKey !== month || item.status !== "late") {
        continue;
      }
      const labelKey = item.label.trim().toLowerCase() || item.id;
      const key = `${labelKey}::${item.monthKey}`;
      if (!lateByKey.has(key)) {
        lateByKey.set(key, item);
      }
    }
  }
  return Array.from(lateByKey.values()).sort((a, b) => {
    const aTime = Date.parse(a.completedAt) || 0;
    const bTime = Date.parse(b.completedAt) || 0;
    return bTime - aTime;
  });
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

/** Prefix for BD9 Excel-row action items (synced from BD8 attachment). */
export const BD9_EXCEL_SOURCE_PREFIX = "xlsx:";

/** Prefix for BD13 mirrors of BD10 late items (local Completed/Late only). */
export const BD13_LATE_SOURCE_PREFIX = "bd13-src:";

/**
 * Sync BD10 late rows onto the BD13 action list.
 * Completing on BD13 only updates these mirrors — upstream BD7/BD8/BD9/BD10
 * statuses stay unchanged.
 */
export function syncBd13LateItemsFromSources(
  list: ClientRequestList,
  sourceLateItems: ReadonlyArray<{
    id: string;
    label: string;
    completedAt?: string;
  }>,
  monthKey: string,
): ClientRequestList {
  const month = monthKey.trim();
  if (!month) {
    return list;
  }
  const sourceById = new Map(
    sourceLateItems
      .map((item) => ({
        id: item.id.trim(),
        label: item.label.trim(),
        completedAt: String(item.completedAt ?? "").trim(),
      }))
      .filter((item) => item.id && item.label)
      .map((item) => [item.id, item] as const),
  );

  const nextItems: ActionWorkflowItem[] = [];
  const seenSourceIds = new Set<string>();

  for (const item of list.actionItems) {
    const isBd13Source = item.sourceRequestId.startsWith(
      BD13_LATE_SOURCE_PREFIX,
    );
    if (item.monthKey !== month) {
      nextItems.push(item);
      continue;
    }
    if (!isBd13Source) {
      if (item.status !== "action") {
        nextItems.push(item);
      }
      continue;
    }
    const sourceId = item.sourceRequestId.slice(BD13_LATE_SOURCE_PREFIX.length);
    const source = sourceById.get(sourceId);
    if (!source) {
      // Upstream no longer late — keep BD13 Completed history only.
      if (item.status === "completed") {
        nextItems.push(item);
      }
      continue;
    }
    seenSourceIds.add(sourceId);
    nextItems.push(
      item.label !== source.label ? { ...item, label: source.label } : item,
    );
  }

  for (const [sourceId, source] of sourceById) {
    if (seenSourceIds.has(sourceId)) {
      continue;
    }
    nextItems.push({
      id: makeItemId("act"),
      label: source.label,
      sourceRequestId: `${BD13_LATE_SOURCE_PREFIX}${sourceId}`,
      monthKey: month,
      status: "late",
      completedAt: source.completedAt || new Date().toISOString(),
    });
  }

  return { ...list, actionItems: nextItems };
}

/**
 * Sync BD8 Excel/CSV data rows onto the BD9 action list for Completed/Late.
 * Preserves resolved statuses; refreshes open labels; drops obsolete open rows.
 */
export function syncBd9ExcelRowsToActionItems(
  list: ClientRequestList,
  rows: ReadonlyArray<{ fingerprint: string; label: string }>,
  monthKey: string,
): ClientRequestList {
  const month = monthKey.trim();
  if (!month) {
    return list;
  }
  const rowByFingerprint = new Map(
    rows
      .map((row) => ({
        fingerprint: row.fingerprint.trim(),
        label: row.label.trim(),
      }))
      .filter((row) => row.fingerprint && row.label)
      .map((row) => [row.fingerprint, row] as const),
  );

  const nextItems: ActionWorkflowItem[] = [];
  const seenFingerprints = new Set<string>();

  for (const item of list.actionItems) {
    const isExcelSource = item.sourceRequestId.startsWith(
      BD9_EXCEL_SOURCE_PREFIX,
    );
    if (item.monthKey !== month) {
      nextItems.push(item);
      continue;
    }
    if (!isExcelSource) {
      // Drop open legacy BD8 question items from the BD9 month list.
      if (item.status !== "action") {
        nextItems.push(item);
      }
      continue;
    }
    const fingerprint = item.sourceRequestId.slice(
      BD9_EXCEL_SOURCE_PREFIX.length,
    );
    const row = rowByFingerprint.get(fingerprint);
    if (!row) {
      if (item.status !== "action") {
        nextItems.push(item);
      }
      continue;
    }
    seenFingerprints.add(fingerprint);
    nextItems.push(
      item.status === "action" && item.label !== row.label
        ? { ...item, label: row.label }
        : item,
    );
  }

  for (const [fingerprint, row] of rowByFingerprint) {
    if (seenFingerprints.has(fingerprint)) {
      continue;
    }
    nextItems.push({
      id: makeItemId("act"),
      label: row.label,
      sourceRequestId: `${BD9_EXCEL_SOURCE_PREFIX}${fingerprint}`,
      monthKey: month,
      status: "action",
      completedAt: "",
    });
  }

  return { ...list, actionItems: nextItems };
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
