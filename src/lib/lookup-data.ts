export type LookupFieldId =
  | "industry"
  | "entityType"
  | "businessActivity"
  | "timeZone"
  | "status"
  | "primaryAssignee"
  | "secondaryAssignee"
  | "bookkeeping"
  | "payroll"
  | "payrollAndSalesTaxes"
  | "deliverables"
  | "yearlyCompliance"
  | "meetingRequirements"
  | "meetingFrequency"
  | "bookkeepingDeadlines"
  | "payrollDeadlines"
  | "payrollTaxesDeadlines"
  | "deliverablesTimelines";

/** Sub-item under a fixed parent field (Industry, Status, etc.). */
export type LookupItem = {
  id: string;
  label: string;
};

export type LookupField = {
  id: LookupFieldId;
  name: string;
  items: LookupItem[];
};

export type LookupDataStore = {
  fields: LookupField[];
};

function compareLabel(a: string, b: string) {
  return a.localeCompare(b, undefined, { sensitivity: "base" });
}

function sortLookupItems(items: LookupItem[]): LookupItem[] {
  return [...items].sort((a, b) => compareLabel(a.label, b.label));
}

function sortLookupFields(fields: LookupField[]): LookupField[] {
  return [...fields]
    .map((field) => ({
      ...field,
      items: sortLookupItems(field.items),
    }))
    .sort((a, b) => compareLabel(a.name, b.name));
}

export function getLookupOptions(
  store: LookupDataStore | null | undefined,
  fieldId: LookupFieldId,
): string[] {
  const labels =
    store?.fields
      .find((field) => field.id === fieldId)
      ?.items.map((item) => item.label) ?? [];
  return [...labels].sort(compareLabel);
}

/** Default Status for new Resource Allocation clients. */
export const DEFAULT_RESOURCE_STATUS = "Onboading";

export function getDefaultStatus(options: string[]): string {
  const match = options.find(
    (option) =>
      option.trim().toLowerCase() === DEFAULT_RESOURCE_STATUS.toLowerCase(),
  );
  // Prefer the exact lookup label when present; otherwise still default to Onboading.
  return match ?? DEFAULT_RESOURCE_STATUS;
}

function makeId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function createLookupItem(label: string): LookupItem {
  return {
    id: makeId("item"),
    label: label.trim(),
  };
}

function seedItems(labels: string[]): LookupItem[] {
  return labels.map((label, index) => ({
    id: `seed-${index + 1}-${label.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`,
    label,
  }));
}

export const LOOKUP_FIELD_META: {
  id: LookupFieldId;
  name: string;
  defaults: string[];
}[] = [
  {
    id: "industry",
    name: "Industry",
    defaults: [
      "Resturant",
      "Installation",
      "Engineering",
      "Online store",
      "Trading",
      "Moving",
      "Service",
      "Educational",
      "Hotel",
      "Construction",
      "Lamination",
      "Environmental services",
      "E commerce",
      "Parlour",
      "Food",
      "Material Handling",
      "Industrial Instrument Trading",
      "Dismantling services",
      "Lamination services",
      "Hotel management",
      "Engineering services",
      "Engineering service",
      "Service provider",
      "Jackets Export",
    ],
  },
  {
    id: "entityType",
    name: "Entity Type",
    defaults: [
      "Business Corporation",
      "C Corpoation",
      "LLC",
      "Partnership",
      "S Corp",
    ],
  },
  {
    id: "businessActivity",
    name: "Business Activity",
    defaults: ["Restaurant/Caterer", "Material Handling"],
  },
  {
    id: "timeZone",
    name: "Time Zone",
    defaults: ["EST", "PST", "Vietnam 2 hours ahead"],
  },
  {
    id: "status",
    name: "Status",
    defaults: ["Ongoing", "Onboading", "Onboarding stage"],
  },
  {
    id: "primaryAssignee",
    name: "Primary Assignee",
    defaults: ["Asad", "Abdul"],
  },
  {
    id: "secondaryAssignee",
    name: "Secondary Assignee",
    defaults: ["Ishaq", "Ahsan", "Ali", "Khizer"],
  },
  {
    id: "bookkeeping",
    name: "Bookkeeping services",
    defaults: [
      "Folder Update",
      "Pos",
      "Quarterly Commission",
      "Inbox",
      "AP and AR Management",
      "Weekly Bookkeeping",
      "Monthly Bookkeeping",
      "Quarterly Bookeeping",
      "Quarterly Budger Reports & Tax Advisory",
    ],
  },
  {
    id: "payroll",
    name: "Payroll",
    defaults: ["NIL", "Bi-Weekly", "Gusto"],
  },
  {
    id: "payrollAndSalesTaxes",
    name: "Payroll and Sales Taxes",
    defaults: [
      "NIL",
      "Monthly Sales Tax",
      "Yearly Sales Tax",
      "Quarterly Sales Tax",
      "Payroll Taxes",
      "Quarterly GST and HST",
      "Yearly GST and HST",
      "WSIB",
    ],
  },
  {
    id: "deliverables",
    name: "Deliverables",
    defaults: [
      "P&L and BS – Monthly",
      "P&L and BS - Quarterly",
      "A/R Report Weekly",
      "A/P Report Weekly",
      "Uncategorized Report Weekly",
      "Payroll Report Biweekly",
      "Quarterly P&L",
      "Monthly P&L",
    ],
  },
  {
    id: "yearlyCompliance",
    name: "Yearly Compliance",
    defaults: [
      "1120",
      "1040",
      "1099",
      "1120S",
      "1065",
      "T1",
      "T2 (Calendar Year, 31 Dec)",
      "T2 (Fiscal Year, 30 June)",
      "T4",
      "T5",
      "NIL",
      "990 Form",
    ],
  },
  {
    id: "meetingRequirements",
    name: "Meeting Requirements",
    defaults: ["NIL", "On demand", "Yes"],
  },
  {
    id: "meetingFrequency",
    name: "Meeting Frequency",
    defaults: ["Weekly", "Bi-Weekly", "Monthly", "Quarterly", "Yearly"],
  },
  {
    id: "bookkeepingDeadlines",
    name: "Bookkeeping Deadlines",
    defaults: ["NIL"],
  },
  {
    id: "payrollDeadlines",
    name: "Payroll Deadlines",
    defaults: ["NIL"],
  },
  {
    id: "payrollTaxesDeadlines",
    name: "Payroll Taxes Deadlines",
    defaults: ["NIL"],
  },
  {
    id: "deliverablesTimelines",
    name: "Deliverables Timelines",
    defaults: ["NIL"],
  },
];

export function createDefaultLookupData(): LookupDataStore {
  return {
    fields: sortLookupFields(
      LOOKUP_FIELD_META.map((field) => ({
        id: field.id,
        name: field.name,
        items: seedItems(field.defaults),
      })),
    ),
  };
}

function flattenLegacyItem(value: unknown): LookupItem[] {
  if (typeof value !== "object" || value === null) {
    return [];
  }

  const raw = value as { id?: unknown; label?: unknown; children?: unknown };
  const label = String(raw.label ?? "").trim();
  const self: LookupItem[] = label
    ? [
        {
          id:
            typeof raw.id === "string" && raw.id
              ? raw.id
              : makeId("item"),
          label,
        },
      ]
    : [];

  const children = Array.isArray(raw.children)
    ? raw.children.flatMap((child) => flattenLegacyItem(child))
    : [];

  return [...self, ...children];
}

export function normalizeLookupData(value: unknown): LookupDataStore {
  const defaults = createDefaultLookupData();
  const incoming =
    typeof value === "object" && value !== null
      ? (value as { fields?: unknown })
      : {};
  const incomingFields = Array.isArray(incoming.fields) ? incoming.fields : [];

  const byId = new Map<string, LookupField>();
  for (const field of incomingFields) {
    if (typeof field !== "object" || field === null) {
      continue;
    }
    const raw = field as Partial<LookupField> & { items?: unknown[] };
    if (typeof raw.id !== "string") {
      continue;
    }
    const meta = LOOKUP_FIELD_META.find((item) => item.id === raw.id);
    if (!meta) {
      continue;
    }

    const items = Array.isArray(raw.items)
      ? raw.items.flatMap((item) => flattenLegacyItem(item))
      : [];

    // Deduplicate by id while preserving order.
    const seen = new Set<string>();
    let uniqueItems = items.filter((item) => {
      if (seen.has(item.id)) {
        return false;
      }
      seen.add(item.id);
      return true;
    });

    // Migrate legacy T2 into the two fiscal variants.
    if (meta.id === "yearlyCompliance") {
      uniqueItems = uniqueItems.filter(
        (item) => item.label.trim().toLowerCase() !== "t2",
      );
    }

    const labels = new Set(
      uniqueItems.map((item) => item.label.trim().toLowerCase()),
    );
    for (const label of meta.defaults) {
      if (!labels.has(label.trim().toLowerCase())) {
        uniqueItems.push(createLookupItem(label));
        labels.add(label.trim().toLowerCase());
      }
    }

    byId.set(raw.id, {
      id: meta.id,
      name: meta.name,
      items: sortLookupItems(uniqueItems),
    });
  }

  return {
    fields: sortLookupFields(
      defaults.fields.map((field) => byId.get(field.id) ?? field),
    ),
  };
}
