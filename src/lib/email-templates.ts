import { stripLegacySignatureClosing } from "@/lib/email-signatures";

export type EmailTemplateId =
  | "template-1"
  | "template-2"
  | "template-3"
  | "template-4"
  | "template-5"
  | "template-6"
  | "template-7"
  | "template-8"
  | "template-9"
  | "template-10"
  | "template-11"
  | "template-12"
  | "template-13"
  | "template-14"
  | "template-15"
  | "template-16"
  | "template-17"
  | "template-18";

export type EmailTemplate = {
  id: string;
  /** Sidebar label, e.g. Template 1 */
  label: string;
  title: string;
  whenSent: string;
  subject: string;
  body: string;
  /** Selected Bookkeeping services lookup values. */
  bookkeepingServices: string[];
  /**
   * Starred templates appear in the main sidebar list.
   * Unstarred are hidden under an expandable section.
   */
  starred: boolean;
};

export type EmailTemplatesStore = {
  templates: EmailTemplate[];
};

const TEMPLATE_META: {
  id: EmailTemplateId;
  title: string;
  whenSent: string;
  subject: string;
  body: string;
}[] = [
  {
    id: "template-1",
    title: "Bank and Credit Card Statements Request",
    whenSent:
      "BD1 of each month (automatic). Lists only accounts not covered by a working bank feed.",
    subject: "[Client Name]: Bank and credit card statements for [Period]",
    body: `Hi [Contact First Name],

We are starting the bookkeeping for [Period] and need the statements for the accounts below:
• [List of bank and credit card accounts, with last 4 digits]

Please upload them to the [portal] by [Due Date]

PDF statements downloaded from your online banking are perfect but a CSV is always welcomed. Please send the full statement, including every page, even if some pages have no transactions.

If an account was closed, or you opened a new account during the month, just let us know in the portal so we can update our records.`,
  },
  {
    id: "template-2",
    title: "Other Statements Request (Loans, Other Institutions, Leases)",
    whenSent:
      "BD1 of each month, or quarterly, depending on how often the lender or lessor issues statements.",
    subject:
      "[Client Name]: Loan, lease and other account statements for [Period]",
    body: `Hi [Contact First Name],

Along with your bank statements, we also need the statements below for [Period] so we can reconcile the balances on your balance sheet:
• [List of loan, line of credit, lease, investment or other accounts]

This usually includes:
• Loan and line of credit statements showing principal, interest and the closing balance
• Equipment or vehicle lease statements and any rental agreements that changed this month
• Statements from any other financial institution, investment or brokerage account held by the business

Please upload them to the [portal] by [Due Date]

If you took out a new loan, signed a new lease, or paid off an existing one during the month, please upload the agreement or payoff letter as well.`,
  },
  {
    id: "template-3",
    title: "Uncategorized Transactions: Questions for the Client",
    whenSent:
      "By BD8, once reconciliations are complete. Transactions are pulled from QBO into the portal question list.",
    subject:
      "[Client Name]: A few transactions we need your help with for [Period]",
    body: `Hi [Contact First Name],

We are working through your [Period] books and have [Number] transactions we could not identify from the bank and card details alone. They total [Total Amount].

You can see the full list in the portal:

[Question List]

For each one, please tell us:
• What it was for (for example, software subscription, materials for a job, owner personal expense)
• Whether it relates to a specific customer, job or project
• A receipt or invoice if you have one (required for anything over [Receipt Threshold])

A short answer is enough. "Lunch with a client" or "Personal, please treat as owner draw" works fine.

Please add attachments where necessary

Please reply by [Due Date] so we can close your books on schedule. Anything still open by then will be recorded as uncategorized and listed in your financial package.`,
  },
  {
    id: "template-4",
    title: "Missing Receipts and Bills",
    whenSent:
      "Alongside the question list (BD8), for transactions above the receipt threshold with no supporting document.",
    subject: "[Client Name]: Receipts needed for [Period]",
    body: `Hi [Contact First Name],

We need receipts or invoices for the following transactions from [Period], as they are over [Receipt Threshold]:
• [List: date, payee, amount]

You can upload them directly against each item in the portal. A photo taken on your phone is fine as long as the date, vendor and amount are readable.

Keeping receipts on file for larger purchases makes things much easier if you are ever asked to support an expense.

Please upload them by [Due Date].`,
  },
  {
    id: "template-5",
    title: "Month-End Coming (Pre-Close Reminder)",
    whenSent: "Last 3 business days of the month (automatic).",
    subject: "[Client Name]: Month-end is coming up, here is what we will need",
    body: `Hi [Contact First Name],

A quick heads-up that [Month] is wrapping up. To keep your books on track, here is what we will be looking for in the first few days of next month:
• [List of statements and reports for this client]
• Receipts for purchases over [Receipt Threshold]
• Any notes on unusual items this month: new loans, large purchases, money put into or taken out of the business

We will send the formal request on [BD1 Date], and the due date will be [Due Date]. If you already have anything ready, you are welcome to upload it to the portal now.`,
  },
  {
    id: "template-6",
    title: "Reminder 1: Missing Documents",
    whenSent:
      "BD7, if any items from the document request are still outstanding (automatic).",
    subject: "Reminder: [Client Name] documents for [Period]",
    body: `Hi [Contact First Name],

Just a friendly reminder that we are still waiting on the following for [Period]:
• [List of outstanding items only]

These were due on [Due Date]. Once we have them we can finish reconciling your accounts and get your financials to you.

You can upload them to the portal

If anything is difficult to get hold of, let us know and we will help.`,
  },
  {
    id: "template-7",
    title: "Reminder 2: Final Reminder for Missing Documents",
    whenSent:
      "BD10, if items are still outstanding. Copied to the business owner / decision-maker (automatic).",
    subject: "Second reminder: [Client Name] documents needed to close [Period]",
    body: `Hi [Contact First Name],

We are following up again as the items below for [Period] are still outstanding:
• [List of outstanding items only]

Without these, we are unable to complete your reconciliations, and your financial statements will be delayed. Our delivery date moves back by the number of days the documents are late, so the current expected date for your financials is [Revised Delivery Date].

Please upload what you can to the portal. If it is easier, reply to this email with a good time and we will give you a quick call to sort it out together.`,
  },
  {
    id: "template-8",
    title: "Reminder: Open Questions",
    whenSent:
      "3 business days after the question list was sent, if any questions are still unanswered (automatic).",
    subject: "Reminder: [Number] open questions on your [Period] books",
    body: `Hi [Contact First Name],

We still have [Number] open questions on your [Period] books, totaling [Total Amount]. You can see and answer them in the portal:

[Question List]

Most of them should only take a few seconds each. If you are unsure about any item, just say so and we will look into it further.

We plan to close your books on [Close Date]. Anything still open by then will be recorded as uncategorized and listed in your financial package.`,
  },
  {
    id: "template-9",
    title: "Timeline Update: Late Documents",
    whenSent:
      "When documents arrive after the due date and the close and delivery dates shift (automatic).",
    subject: "[Client Name]: Updated delivery date for your [Period] financials",
    body: `Hi [Contact First Name],

Thank you for sending your documents for [Period]. Since they arrived [Number] business days after the due date, your financials will now be ready by [Revised Delivery Date] instead of [Original Delivery Date].

We will keep you posted if anything else comes up. Going forward, having everything in by [Standard Due Day] each month lets us deliver your numbers on the usual schedule.`,
  },
  {
    id: "template-10",
    title: "Bank Feed Disconnected",
    whenSent:
      "When the portal detects that a QBO bank feed has stopped updating for more than a few days.",
    subject: "Action needed: [Bank Name] connection to QuickBooks has stopped",
    body: `Hi [Contact First Name],

The connection between your [Bank Name] account ending in [Last 4 Digits] and QuickBooks Online stopped updating on [Date]. This usually happens when a password is changed or the bank asks for the connection to be re-approved.

To fix it, please log in to QuickBooks, go to Transactions, then Bank transactions, and follow the prompt to reconnect the account. It normally takes a couple of minutes.

If you would rather not reconnect, please upload the [Period] statement to the portal instead

Let us know if you run into any trouble and we can walk you through it.`,
  },
  {
    id: "template-11",
    title: "Your Financial Statements Are Ready",
    whenSent: "When the monthly package is released to the client portal after review.",
    subject: "[Client Name]: Your financial statements for [Period] are ready",
    body: `Hi [Contact First Name],

Your financial statements for [Period] are ready and available in the portal

The package includes:
• Profit and Loss (month and year to date)
• Balance Sheet
• Accounts receivable and accounts payable aging

A few highlights for the month: [if applicable]
• [Short commentary point 1]
• [Short commentary point 2]
• [Short commentary point 3]

[If applicable] There are [Number] open items still waiting on your input, listed at the end of the package.

If you have any questions, reply to this email or book a time with us here: [Booking Link]`,
  },
  {
    id: "template-12",
    title: "Review Meeting: Invitation and Agenda",
    whenSent: "2 business days before the scheduled review meeting (automatic).",
    subject: "[Client Name]: Agenda for our [Period] review on [Meeting Date]",
    body: `Hi [Contact First Name],

Looking forward to our call on [Meeting Date] at [Meeting Time, client time zone]. Here is the link to join: [Meeting Link]

Agenda:
• Quick walk-through of the [Period] results
• Cash position and anything coming up
• Open questions: [Number] items, listed in the portal
• Anything you would like to discuss

Your financial statements are in the portal. If you have a chance to look before the call, we can spend more of the time on what matters to you.

If the time no longer works, you can reschedule here: [Booking Link]`,
  },
  {
    id: "template-13",
    title: "Review Meeting: Recap and Action Items",
    whenSent: "Same day as the review meeting.",
    subject: "[Client Name]: Recap of our [Period] review",
    body: `Hi [Contact First Name],

Thanks for your time today. Here is a quick summary of what we covered.

Key points / Minutes of the Meeting:
• [Point 1]
• [Point 2]

Action items:
• [Action, owner, due date]
• [Action, owner, due date]

If we missed anything, just reply to this email. Our next review is planned for [Next Meeting Date].`,
  },
  {
    id: "template-14",
    title: "New Client Welcome and Onboarding",
    whenSent: "When a new bookkeeping client is set up in the portal.",
    subject: "Welcome to [Firm Name]: getting your bookkeeping set up",
    body: `Hi [Contact First Name],

Welcome aboard. We are glad to be working with you. Your portal account is ready, and you can log in to the portal

To get started, we need the following:
• Accountant (user) access to your QuickBooks Online company
• A list of all bank, credit card, loan and lease accounts used by the business
• Your most recent bank and credit card statements for each account
• Payroll provider details and access to payroll reports, if applicable
• Your prior year financial statements or last tax return, if available
• Any loan or lease agreements currently in place

Your main contact here will be [Staff Name], and [Manager Name] will review your work each month.

How the month will work:
• Around the 1st, we will request your statements and documents
• Documents are due by [Standard Due Day]
• Around the [Question Day], we will send any questions about transactions
• Your financial statements will be ready by [Standard Delivery Day]

Please upload the items above by [Due Date]. If you have any questions, just reply here.`,
  },
  {
    id: "template-15",
    title: "Year-End Bookkeeping Documents",
    whenSent:
      "In the last month of the client's fiscal year, with a reminder after year-end.",
    subject: "[Client Name]: Year-end bookkeeping items for [Fiscal Year]",
    body: `Hi [Contact First Name],

As [Fiscal Year] comes to a close, we will be finalizing your books for the full year. In addition to your regular monthly documents, please send the following:
• December (or year-end month) statements for all bank, credit card, loan and lease accounts
• Year-end loan statements showing the interest paid for the year
• Details of any equipment, vehicles or other assets purchased or sold during the year
• A list of customers who owed you money at year-end, and any amounts you do not expect to collect
• A list of unpaid bills at year-end
• Inventory count at year-end, if applicable
• Details of any money put into or taken out of the business by the owners

Please upload these to the [portal] by [Due Date]

Getting these in early means your final year-end numbers will be ready by [Year-End Delivery Date].`,
  },
  {
    id: "template-16",
    title: "Pakistan Holiday Notice",
    whenSent:
      "Last weekday before each Pakistan holiday series (9:00 AM Pakistan time; skips Sat/Sun). Multi-day holidays like Eid Day 1–3 share one email. Sent to the whole team in User Management from ayesha@fireflyaccountancy.com or sana@epsfresno.com.",
    subject: "Reminder: Upcoming Pakistan holiday — [Holiday Name]",
    body: `Hi team,

This is a reminder of an upcoming Pakistan holiday:

• [Holiday Name] — [Holiday Date]

Please plan your work accordingly.

Thank you.`,
  },
  {
    id: "template-17",
    title: "Pakistan Holiday Notice (Clients)",
    whenSent:
      "15 calendar days before each Pakistan holiday series (weekends count), then shifted to the previous weekday if that date is Sat/Sun (9:00 AM in each client's time zone). Multi-day holidays share one email. Sent to all Client Profile emails from ayesha@fireflyaccountancy.com or sana@epsfresno.com.",
    subject: "Notice: Upcoming Pakistan holiday — [Holiday Name]",
    body: `Hi,

This is a notice that our Pakistan office will be closed for an upcoming holiday:

• [Holiday Name] — [Holiday Date]

Please plan any document requests or replies accordingly. We will respond as soon as we return.

Thank you.`,
  },
  {
    id: "template-18",
    title: "End-of-Month Team Reminder",
    whenSent:
      "Last 3 / Last 2 / Last 1 BD in Client Management → Team button. One template for all three sidebar days.",
    subject: "Team reminder: [Last BD Label] — complete pending work",
    body: `Hi team,

This is a reminder that we are in the [Last BD Label].

Please complete all pending Client Management work before month-end:
• Outstanding document requests and late items
• Open questions and follow-ups
• Reviews, packages, and status checks still in progress

Finish what you can today so nothing carries into next month unnecessarily.

Thank you.`,
  },
];

const SEED_TEMPLATE_IDS = TEMPLATE_META.map((item) => item.id);

function templateNumber(id: string): number {
  const match = id.match(/^template-(\d+)$/);
  return match ? Number(match[1]) : 0;
}

/** Number from sidebar label, e.g. "Template 11 (Client)" → 11. */
export function emailTemplateLabelNumber(template: EmailTemplate): number {
  const fromLabel = template.label.trim().match(/^Template\s+(\d+)\b/i);
  if (fromLabel) {
    return Number(fromLabel[1]);
  }
  const fromId = templateNumber(template.id);
  return fromId > 0 ? fromId : Number.MAX_SAFE_INTEGER;
}

/** Sort templates 1, 2, 3 … 10, 11 by their label number. */
export function compareEmailTemplatesByLabelNumber(
  a: EmailTemplate,
  b: EmailTemplate,
): number {
  const diff = emailTemplateLabelNumber(a) - emailTemplateLabelNumber(b);
  if (diff !== 0) {
    return diff;
  }
  return a.label.localeCompare(b.label, undefined, { sensitivity: "base" });
}

function isSeedTemplateId(id: string): id is EmailTemplateId {
  return SEED_TEMPLATE_IDS.includes(id as EmailTemplateId);
}

function makeTemplateId() {
  return `template-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export function suggestNextTemplateLabel(templates: EmailTemplate[]): string {
  let max = 0;
  for (const template of templates) {
    const match = template.label.trim().match(/^Template\s+(\d+)$/i);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `Template ${max + 1}`;
}

function normalizeBookkeepingServices(value: unknown): string[] {
  if (Array.isArray(value)) {
    const seen = new Set<string>();
    const next: string[] = [];
    for (const item of value) {
      const label = String(item ?? "").trim();
      if (!label) {
        continue;
      }
      const key = label.toLowerCase();
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      next.push(label);
    }
    return next;
  }

  // Legacy single category string (e.g. "bookkeeping") is not a service list.
  return [];
}

export function createEmailTemplate(
  input: Omit<EmailTemplate, "id"> & { starred?: boolean },
): EmailTemplate {
  return {
    id: makeTemplateId(),
    label: input.label.trim(),
    title: input.title.trim(),
    whenSent: input.whenSent.trim(),
    subject: input.subject.trim(),
    body: stripLegacySignatureClosing(input.body),
    bookkeepingServices: normalizeBookkeepingServices(
      input.bookkeepingServices,
    ),
    starred: input.starred !== false,
  };
}

function normalizeOneTemplate(
  raw: Partial<EmailTemplate> & { category?: unknown },
  fallback?: EmailTemplate,
): EmailTemplate | null {
  const id = String(raw.id ?? fallback?.id ?? "").trim();
  const title = String(raw.title ?? fallback?.title ?? "").trim();
  if (!id || !title) {
    return null;
  }
  const number = templateNumber(id);
  const bookkeepingServices =
    raw.bookkeepingServices !== undefined
      ? normalizeBookkeepingServices(raw.bookkeepingServices)
      : (fallback?.bookkeepingServices ?? []);
  // Missing starred → keep visible (starred) so existing templates stay in the list.
  const starred =
    raw.starred === false
      ? false
      : raw.starred === true
        ? true
        : (fallback?.starred ?? true);
  return {
    id,
    label: String(
      raw.label ?? fallback?.label ?? (number ? `Template ${number}` : "Template"),
    ).trim(),
    title,
    whenSent: String(raw.whenSent ?? fallback?.whenSent ?? "").trim(),
    subject: String(raw.subject ?? fallback?.subject ?? "").trim(),
    body: stripLegacySignatureClosing(
      String(raw.body ?? fallback?.body ?? ""),
    ),
    bookkeepingServices,
    starred,
  };
}

/** Calendar month and year for a date, e.g. "October 2026". */
export function formatCurrentPeriod(date: Date = new Date()): string {
  return date.toLocaleDateString("en-US", {
    month: "long",
    year: "numeric",
  });
}

/** Month name only, e.g. "October". */
export function formatPeriodMonth(date: Date = new Date()): string {
  return date.toLocaleDateString("en-US", { month: "long" });
}

/** Year only, e.g. "2026". */
export function formatPeriodYear(date: Date = new Date()): string {
  return String(date.getFullYear());
}

/**
 * Replace [Period] / [Month] / [Year] using the given date
 * (Client Management active / demo month when provided).
 */
export function applyPeriodPlaceholder(
  text: string,
  date: Date = new Date(),
): string {
  const period = formatCurrentPeriod(date);
  const month = formatPeriodMonth(date);
  const year = formatPeriodYear(date);
  return text
    .replace(/\[Period\]/gi, period)
    .replace(/\[Month\]/gi, month)
    .replace(/\[Year\]/gi, year);
}

/** Subject and body only (header fields shown separately). */
export function formatEmailTemplateDisplay(template: EmailTemplate): string {
  const lines = [
    `Subject: ${applyPeriodPlaceholder(template.subject)}`,
    applyPeriodPlaceholder(template.body),
  ].filter(Boolean);
  return lines.join("\n\n");
}

export function createDefaultEmailTemplates(): EmailTemplatesStore {
  return {
    templates: TEMPLATE_META.map((item) => {
      const number = templateNumber(item.id);
      return {
        id: item.id,
        label: `Template ${number}`,
        title: item.title,
        whenSent: item.whenSent,
        subject: item.subject,
        body: stripLegacySignatureClosing(item.body),
        bookkeepingServices: [],
        starred: true,
      };
    }),
  };
}

export function normalizeEmailTemplates(value: unknown): EmailTemplatesStore {
  const defaults = createDefaultEmailTemplates();
  const incoming =
    typeof value === "object" && value !== null
      ? (value as { templates?: unknown })
      : {};
  const list = Array.isArray(incoming.templates) ? incoming.templates : [];

  const seedById = new Map<string, EmailTemplate>();
  const customById = new Map<string, EmailTemplate>();

  for (const item of list) {
    if (typeof item !== "object" || item === null) {
      continue;
    }
    const raw = item as Partial<EmailTemplate>;
    const id = String(raw.id ?? "").trim();
    if (!id) {
      continue;
    }
    const fallback = defaults.templates.find((entry) => entry.id === id);
    const normalized = normalizeOneTemplate(raw, fallback);
    if (!normalized) {
      continue;
    }
    if (isSeedTemplateId(id)) {
      seedById.set(id, normalized);
    } else {
      customById.set(id, normalized);
    }
  }

  const mergedSeeds = defaults.templates.map(
    (template) => seedById.get(template.id) ?? template,
  );
  const customs = [...customById.values()].filter(
    (template) => !isSeedTemplateId(template.id),
  );

  return {
    templates: [...mergedSeeds, ...customs],
  };
}
