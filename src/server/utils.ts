import fs from "node:fs/promises";
import path from "node:path";
import * as XLSX from "xlsx";

const XLSXLib: typeof XLSX = (XLSX as unknown as { default?: typeof XLSX }).default ?? XLSX;

export async function ensureDir(dir: string) {
  await fs.mkdir(dir, { recursive: true });
}

export function safePathSegment(value: string) {
  return value.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").trim() || "未命名";
}

export function timestampForPath(date = new Date()) {
  const pad = (n: number) => `${n}`.padStart(2, "0");
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate())
  ].join("-") + `_${pad(date.getHours())}-${pad(date.getMinutes())}-${pad(date.getSeconds())}`;
}

function normalizeHeader(value: unknown) {
  return String(value ?? "").trim();
}

function isEmptyCell(value: unknown) {
  return value === null || value === undefined || String(value).trim() === "";
}

function cellDisplayValue(cell: XLSX.CellObject | undefined) {
  if (!cell) return "";
  if (cell.v instanceof Date) return cell.v;
  return cell.w ?? cell.v ?? "";
}

export function readWorkbookRows(filePath: string, selectedFields?: string[]) {
  const workbook = XLSXLib.readFile(filePath, { cellDates: true, raw: false });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) return [];
  const sheet = workbook.Sheets[firstSheet];
  if (!selectedFields?.length) {
    return XLSXLib.utils.sheet_to_json<Record<string, unknown>>(sheet, {
      defval: "",
      raw: false
    });
  }

  const range = XLSXLib.utils.decode_range(sheet["!ref"] ?? "A1:A1");
  const headerIndexes = new Map<string, number>();
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const header = normalizeHeader(cellDisplayValue(sheet[XLSXLib.utils.encode_cell({ r: range.s.r, c: col })]));
    if (header && !headerIndexes.has(header)) headerIndexes.set(header, col);
  }

  const rows: Record<string, unknown>[] = [];
  const selectedIndexes = selectedFields.map((field) => headerIndexes.get(field));
  for (let rowIndex = range.s.r + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const row: Record<string, unknown> = {};
    let hasValue = false;
    selectedFields.forEach((field, fieldIndex) => {
      const col = selectedIndexes[fieldIndex];
      const value = col === undefined ? "" : cellDisplayValue(sheet[XLSXLib.utils.encode_cell({ r: rowIndex, c: col })]);
      row[field] = value;
      if (!isEmptyCell(value)) hasValue = true;
    });
    if (hasValue) rows.push(row);
  }
  return rows;
}

export function validateWorkbookFields(filePath: string, requiredFields: string[]) {
  const workbook = XLSXLib.readFile(filePath, { cellDates: true, raw: false, sheetRows: 1 });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) return requiredFields;
  const rows = XLSXLib.utils.sheet_to_json<unknown[]>(workbook.Sheets[firstSheet], {
    header: 1,
    defval: "",
    raw: false
  });
  const headers = new Set((rows[0] ?? []).map(normalizeHeader));
  return requiredFields.filter((field) => !headers.has(field));
}

export async function writeJson(filePath: string, data: unknown) {
  await ensureDir(path.dirname(filePath));
  await fs.writeFile(filePath, JSON.stringify(data, null, 2), "utf8");
}

export async function readJson<T>(filePath: string, fallback: T): Promise<T> {
  try {
    const content = await fs.readFile(filePath, "utf8");
    return JSON.parse(content) as T;
  } catch {
    return fallback;
  }
}

export function countBy<T>(items: T[], key: (item: T) => string) {
  return items.reduce<Record<string, number>>((acc, item) => {
    const value = key(item);
    acc[value] = (acc[value] ?? 0) + 1;
    return acc;
  }, {});
}
