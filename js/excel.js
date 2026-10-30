/* global XLSX */
/**
 * Helpers around SheetJS, which index.html loads as the global `XLSX`.
 * Shared by the Data Cleaner and the Report Merger, so keep this file identical across those repos.
 *
 * A cell value is one of: a string, a number, a boolean, or a date `{ date: true, serial, z }`
 * where `serial` is the Excel day number and `z` is its number format.
 */

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;
const MAX_COLUMN_WIDTH = 50;
const MIN_COLUMN_WIDTH = 8;

export const ISO_DATE_FORMAT = 'yyyy-mm-dd';

/** Excel serial number for a calendar date (correct for dates after 1 March 1900). */
export const ymdToSerial = (year, month, day) => (Date.UTC(year, month - 1, day) - EXCEL_EPOCH) / MS_PER_DAY;

export const isDate = (value) => value !== null && typeof value === 'object' && value.date === true;

/** Text a person would see for a value. Dates use their own number format. */
export function show(value) {
  if (value === '' || value == null) return '';
  if (isDate(value)) return XLSX.SSF.format(value.z, value.serial);
  if (typeof value === 'number') return String(Math.round(value * 1e6) / 1e6);
  return String(value);
}

export const isBlank = (value) => show(value).trim() === '';

/** Turns a SheetJS cell into a plain value, keeping real dates as dates. */
export function cellValue(cell) {
  if (!cell || cell.v == null) return '';
  if (cell.t === 'n') return cell.z && XLSX.SSF.is_date(cell.z) ? { date: true, serial: cell.v, z: cell.z } : cell.v;
  if (cell.t === 'b') return cell.v;
  if (cell.t === 'e') return cell.w || '#ERROR';
  return String(cell.v);
}

/**
 * Reads a worksheet into rows.
 * @returns {{ n: number, cells: Array }[]} `n` is the row number as Excel shows it.
 */
export function sheetToGrid(sheet) {
  if (!sheet || !sheet['!ref']) return [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const grid = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const cells = [];
    for (let c = range.s.c; c <= range.e.c; c++) cells.push(cellValue(sheet[XLSX.utils.encode_cell({ r, c })]));
    grid.push({ n: r + 1, cells });
  }
  return grid;
}

/** Index of the first row with anything in it, which is treated as the header row. */
export const findHeaderRow = (grid) => grid.findIndex((row) => row.cells.some((value) => !isBlank(value)));

/** Header names from a row, tidied, with a fallback name for blank headers. */
export const headerNames = (cells) => cells.map((value, i) => show(value).replace(/\s+/g, ' ').trim() || `Column ${i + 1}`);

/**
 * Opens an .xlsx, .xls or .csv file.
 * CSV is read as plain text, so nothing is guessed before the tool's own rules run.
 */
export async function readWorkbook(file) {
  if (/\.csv$/i.test(file.name)) {
    return { workbook: XLSX.read(await file.text(), { type: 'string', raw: true }), isCsv: true };
  }
  return { workbook: XLSX.read(await file.arrayBuffer(), { type: 'array', cellNF: true }), isCsv: false };
}

/** Builds a worksheet from rows of values, keeping date formats and sizing the columns to fit. */
export function sheetFromRows(rows) {
  const sheet = XLSX.utils.aoa_to_sheet(rows.map((row) => row.map((value) => (isDate(value) ? value.serial : value))));
  // One pass sets date formats and measures column widths. A plain loop, not Math.max(...rows),
  // because spreading a very large sheet into arguments can overflow the call stack.
  const widths = rows[0].map(() => MIN_COLUMN_WIDTH);
  rows.forEach((row, r) => row.forEach((value, c) => {
    if (isDate(value)) sheet[XLSX.utils.encode_cell({ r, c })].z = value.z;
    if (c < widths.length) widths[c] = Math.max(widths[c], Math.min(MAX_COLUMN_WIDTH, show(value).length + 2));
  }));
  sheet['!cols'] = widths.map((wch) => ({ wch }));
  return sheet;
}

/** Builds a workbook from `[sheetName, rows]` pairs. */
export function workbookFrom(sheets) {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of sheets) XLSX.utils.book_append_sheet(workbook, sheetFromRows(rows), name);
  return workbook;
}

/** Saves `[sheetName, rows]` pairs as an .xlsx download. */
export function downloadWorkbook(fileName, sheets) {
  XLSX.writeFile(workbookFrom(sheets), fileName);
}
