/**
 * Standard Client Management schedule slots (sidebar BD list).
 * Used by monthly repeated-task refresh.
 */
export type ClientManagementBdSlot = {
  id: string;
  label: string;
  /** Default email / request template summary for this BD. */
  defaultTemplate: string;
  /** Due-date rule shown in Client Management emails. */
  deadlineRule: string;
};

export const CLIENT_MANAGEMENT_BD_SLOTS: ClientManagementBdSlot[] = [
  {
    id: "bd1",
    label: "BD1",
    defaultTemplate: "Template 1 or Template 2 (by client BD1 Template)",
    deadlineRule: "BD4",
  },
  {
    id: "bd5",
    label: "BD5",
    defaultTemplate: "Review / follow-up action items",
    deadlineRule: "—",
  },
  {
    id: "bd7",
    label: "BD7",
    defaultTemplate: "Late items reminder (BD1 lists)",
    deadlineRule: "BD8 (or BD7 if weekend between)",
  },
  {
    id: "bd8",
    label: "BD8",
    defaultTemplate: "Template 3 — Questions list",
    deadlineRule: "BD8",
  },
  {
    id: "bd9",
    label: "BD9",
    defaultTemplate: "Template 8",
    deadlineRule: "BD9",
  },
  {
    id: "bd10",
    label: "BD10",
    defaultTemplate: "Template 7",
    deadlineRule: "BD10",
  },
  {
    id: "bd15",
    label: "BD15",
    defaultTemplate: "Template 11 — Monthly package",
    deadlineRule: "—",
  },
  {
    id: "bd18",
    label: "BD18",
    defaultTemplate: "Template 11 — Status check",
    deadlineRule: "—",
  },
  {
    id: "last-3-bd",
    label: "Last 3 BD",
    defaultTemplate: "Client: Template 5 · Team: Template 18",
    deadlineRule: "—",
  },
  {
    id: "last-2-bd",
    label: "Last 2 BD",
    defaultTemplate: "Client: Template 12 · Team: Template 18",
    deadlineRule: "—",
  },
  {
    id: "last-1-bd",
    label: "Last 1 BD",
    defaultTemplate: "Client: End-of-month follow-up · Team: Template 18",
    deadlineRule: "—",
  },
];

export function findClientManagementBdSlot(
  id: string,
): ClientManagementBdSlot | null {
  return CLIENT_MANAGEMENT_BD_SLOTS.find((slot) => slot.id === id) ?? null;
}
