import * as XLSX from "xlsx";

/**
 * BD8 question-list spreadsheet stats:
 * [Number] = filled rows − 1 (header)
 * [Total Amount] = sum of the Amount column
 */

export type Bd8AttachmentSpreadsheetStats = {
  dataRowCount: number;
  /** Sum of Amount column values; null if the column is missing. */
  totalAmount: number | null;
};

function decodeBase64ToUint8Array(contentBytes: string): Uint8Array {
  const binary =
    typeof atob === "function"
      ? atob(contentBytes)
      : Buffer.from(contentBytes, "base64").toString("binary");
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function decodeBase64ToUtf8(contentBytes: string): string {
  const bytes = decodeBase64ToUint8Array(contentBytes);
  if (typeof TextDecoder !== "undefined") {
    return new TextDecoder("utf-8").decode(bytes);
  }
  return Buffer.from(bytes).toString("utf-8");
}

function isCsvFile(name: string, contentType: string): boolean {
  const lower = name.trim().toLowerCase();
  const type = contentType.trim().toLowerCase();
  return (
    lower.endsWith(".csv") ||
    type.includes("csv") ||
    type === "text/plain"
  );
}

function cellText(cell: unknown): string {
  if (cell == null) {
    return "";
  }
  if (typeof cell === "string") {
    return cell.trim();
  }
  if (typeof cell === "number" || typeof cell === "boolean") {
    return String(cell);
  }
  return String(cell).trim();
}

function isDateLikeHeader(header: string): boolean {
  const normalized = header.toLowerCase().replace(/[\s_]+/g, " ").trim();
  if (!normalized) {
    return false;
  }
  return (
    normalized === "date" ||
    normalized.startsWith("date ") ||
    normalized.endsWith(" date") ||
    normalized.includes(" date ") ||
    normalized === "datetime" ||
    normalized === "date time" ||
    normalized === "transaction date" ||
    normalized === "posted date"
  );
}

function formatJsDate(value: Date): string {
  const month = value.getMonth() + 1;
  const day = value.getDate();
  const year = value.getFullYear();
  const hours = value.getHours();
  const minutes = value.getMinutes();
  const seconds = value.getSeconds();
  const hasTime = hours !== 0 || minutes !== 0 || seconds !== 0;
  const datePart = `${month}/${day}/${year}`;
  if (!hasTime) {
    return datePart;
  }
  const period = hours >= 12 ? "PM" : "AM";
  const hour12 = hours % 12 || 12;
  const mm = String(minutes).padStart(2, "0");
  return `${datePart} ${hour12}:${mm} ${period}`;
}

/** Excel stores dates as day serials (e.g. 46032 → 1/10/2026). */
function formatExcelSerialDate(serial: number): string | null {
  if (!Number.isFinite(serial) || serial < 1 || serial >= 100000) {
    return null;
  }
  try {
    const hasTime = Math.abs(serial % 1) > 1e-10;
    return String(
      XLSX.SSF.format(hasTime ? "m/d/yyyy h:mm AM/PM" : "m/d/yyyy", serial),
    ).trim();
  } catch {
    return null;
  }
}

function formatCellValue(cell: unknown, header = ""): string {
  if (cell instanceof Date && !Number.isNaN(cell.getTime())) {
    return formatJsDate(cell);
  }
  if (
    typeof cell === "number" &&
    Number.isFinite(cell) &&
    isDateLikeHeader(header)
  ) {
    const formatted = formatExcelSerialDate(cell);
    if (formatted) {
      return formatted;
    }
  }
  return cellText(cell);
}

function rowHasValue(cells: unknown[]): boolean {
  return cells.some((cell) => cellText(cell).length > 0);
}

/** Parse "$1,234.56", "(100)", "1 234,56" style values into a number. */
export function parseSpreadsheetAmount(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  let text = cellText(value);
  if (!text) {
    return null;
  }
  text = text.replace(/\s/g, "");
  let negative = false;
  if (/^\(.*\)$/.test(text)) {
    negative = true;
    text = text.slice(1, -1);
  }
  if (text.startsWith("-")) {
    negative = true;
    text = text.slice(1);
  }
  text = text.replace(/[$€£¥]/g, "");
  // European 1.234,56 → 1234.56
  if (/^\d{1,3}(\.\d{3})+,\d+$/.test(text)) {
    text = text.replace(/\./g, "").replace(",", ".");
  } else if (text.includes(",") && !text.includes(".")) {
    text = text.replace(",", ".");
  } else {
    text = text.replace(/,/g, "");
  }
  const parsed = Number(text);
  if (!Number.isFinite(parsed)) {
    return null;
  }
  return negative ? -parsed : parsed;
}

function normalizeHeader(value: unknown): string {
  return cellText(value).toLowerCase().replace(/[\s_]+/g, " ");
}

function findAmountColumnIndex(headerRow: unknown[]): number {
  const exact = headerRow.findIndex(
    (cell) => normalizeHeader(cell) === "amount",
  );
  if (exact >= 0) {
    return exact;
  }
  return headerRow.findIndex((cell) => {
    const header = normalizeHeader(cell);
    return header === "total amount" || header.endsWith(" amount");
  });
}

function splitCsvLine(line: string): string[] {
  const cells: string[] = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i]!;
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (ch === "," && !inQuotes) {
      cells.push(current);
      current = "";
      continue;
    }
    current += ch;
  }
  cells.push(current);
  return cells;
}

export type Bd8AttachmentTableRow = {
  fingerprint: string;
  cells: string[];
  label: string;
};

export type Bd8AttachmentTable = {
  headers: string[];
  rows: Bd8AttachmentTableRow[];
};

function fingerprintCells(cells: string[]): string {
  const raw = cells.map((cell) => cell.trim().toLowerCase()).join("\u0001");
  let hash = 0;
  for (let i = 0; i < raw.length; i += 1) {
    hash = (hash << 5) - hash + raw.charCodeAt(i);
    hash |= 0;
  }
  return `r${Math.abs(hash).toString(36)}-${raw.length}`;
}

function formatRowLabel(headers: string[], cells: string[]): string {
  const parts: string[] = [];
  const width = Math.max(headers.length, cells.length);
  for (let i = 0; i < width; i += 1) {
    const header = (headers[i] ?? "").trim();
    const value = (cells[i] ?? "").trim();
    if (!header && !value) {
      continue;
    }
    parts.push(header ? `${header}: ${value}` : value);
  }
  return parts.join(" · ");
}

function tableFromRows(rows: unknown[][]): Bd8AttachmentTable {
  const nonEmpty = rows.filter((row) => Array.isArray(row) && rowHasValue(row));
  if (nonEmpty.length === 0) {
    return { headers: [], rows: [] };
  }
  const headerCells = nonEmpty[0]!.map((cell) => cellText(cell));
  const dataRows = nonEmpty.slice(1);
  const parsedRows: Bd8AttachmentTableRow[] = [];
  for (const row of dataRows) {
    const cells = headerCells.map((header, index) =>
      formatCellValue(row[index], header),
    );
    if (!rowHasValue(cells)) {
      continue;
    }
    parsedRows.push({
      fingerprint: fingerprintCells(cells),
      cells,
      label: formatRowLabel(headerCells, cells),
    });
  }
  return { headers: headerCells, rows: parsedRows };
}

function statsFromRows(rows: unknown[][]): Bd8AttachmentSpreadsheetStats {
  const table = tableFromRows(rows);
  if (table.rows.length === 0) {
    return { dataRowCount: 0, totalAmount: null };
  }
  const amountIndex = findAmountColumnIndex(table.headers);
  let totalAmount: number | null = null;
  if (amountIndex >= 0) {
    totalAmount = 0;
    for (const row of table.rows) {
      const amount = parseSpreadsheetAmount(row.cells[amountIndex]);
      if (amount != null) {
        totalAmount += amount;
      }
    }
  }
  return {
    dataRowCount: table.rows.length,
    totalAmount,
  };
}

function loadSpreadsheetRows(
  name: string,
  contentType: string,
  contentBytes: string,
): unknown[][] {
  return isCsvFile(name, contentType)
    ? csvRows(contentBytes)
    : excelRows(contentBytes);
}

function csvRows(contentBytes: string): unknown[][] {
  const text = decodeBase64ToUtf8(contentBytes).replace(/^\uFEFF/, "");
  return text
    .split(/\r\n|\n|\r/)
    .filter((line) => line.trim().length > 0)
    .map((line) => splitCsvLine(line));
}

function excelRows(contentBytes: string): unknown[][] {
  const bytes = decodeBase64ToUint8Array(contentBytes);
  const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    return [];
  }
  const sheet = workbook.Sheets[firstSheetName];
  if (!sheet) {
    return [];
  }
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    raw: true,
    blankrows: false,
    defval: "",
  });
}

export function parseBd8AttachmentTable(
  name: string,
  contentType: string,
  contentBytes: string,
): Bd8AttachmentTable {
  const bytes = contentBytes.trim();
  if (!bytes) {
    return { headers: [], rows: [] };
  }
  try {
    return tableFromRows(loadSpreadsheetRows(name, contentType, bytes));
  } catch {
    return { headers: [], rows: [] };
  }
}

export function getBd8AttachmentSpreadsheetStats(
  name: string,
  contentType: string,
  contentBytes: string,
): Bd8AttachmentSpreadsheetStats {
  const bytes = contentBytes.trim();
  if (!bytes) {
    return { dataRowCount: 0, totalAmount: null };
  }
  try {
    return statsFromRows(loadSpreadsheetRows(name, contentType, bytes));
  } catch {
    return { dataRowCount: 0, totalAmount: null };
  }
}

/** Filled spreadsheet rows excluding the header (never negative). */
export function countBd8AttachmentDataRows(
  name: string,
  contentType: string,
  contentBytes: string,
): number {
  return getBd8AttachmentSpreadsheetStats(name, contentType, contentBytes)
    .dataRowCount;
}

/** Format a summed Amount for [Total Amount] in the BD8 email. */
export function formatBd8TotalAmount(total: number): string {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(total);
}

/**
 * Pull Amount from a BD8/BD9 row label like
 * "DATE: … · DESCRIPTION: … · AMOUNT: 200".
 */
export function parseAmountFromActionLabel(label: string): number | null {
  const match = label.match(/(?:^|[·•|]\s*)Amount\s*:\s*([^·•|]+)/i);
  if (!match) {
    return null;
  }
  return parseSpreadsheetAmount(match[1]);
}

/** Sum Amount values from Late action labels (BD13 [Total Amount]). */
export function sumLateActionLabelAmounts(
  items: ReadonlyArray<{ label: string; status?: string }>,
): number | null {
  let total = 0;
  let found = false;
  for (const item of items) {
    if (item.status && item.status !== "late") {
      continue;
    }
    const amount = parseAmountFromActionLabel(item.label);
    if (amount == null) {
      continue;
    }
    total += amount;
    found = true;
  }
  return found ? total : null;
}
