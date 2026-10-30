/**
 * The merging rules. Pure functions with no page code, so they can be read and tested on their own.
 */
import { findHeaderRow, headerNames, isBlank, sheetToGrid, show } from './excel.js';
import { parseNumber } from './parse.js';

// A column counts as numeric when at least this share of its filled cells are numbers.
const NUMERIC_SHARE = 0.8;

const round2 = (n) => Math.round(n * 100) / 100;

/** Key used to match columns across files: "Sales (RM)", "SALES RM" and "sales(rm)" all become "salesrm". */
export const columnKey = (header) => header.toLowerCase().replace(/[^a-z0-9]/g, '');

/** The number in a cell, reading text like "RM 1,250" too. Undefined when the cell is not a number. */
export function numberIn(value) {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') return parseNumber(value)?.value;
  return undefined;
}

/** A CSV has no number type, so plain numbers in a CSV are turned into real numbers. */
function fromCsvText(text) {
  const number = parseNumber(text);
  return number && !number.formatted ? number.value : text;
}

function readPart(file, sheetName) {
  const grid = sheetToGrid(file.workbook.Sheets[sheetName]);
  const headerIndex = findHeaderRow(grid);
  if (headerIndex < 0) return null;
  const rows = grid.slice(headerIndex + 1)
    .map(({ cells }) => (file.isCsv ? cells.map((v) => (typeof v === 'string' ? fromCsvText(v) : v)) : cells))
    .filter((cells) => cells.some((v) => !isBlank(v)));
  return { file: file.name, sheet: sheetName, headers: headerNames(grid[headerIndex].cells), rows };
}

// Reading a sheet is the slow part, so each file's parts are read once and reused.
const partsCache = new WeakMap();

/**
 * The sheets of one file as parts: { file, sheet, headers, rows }.
 * @param {{ name: string, workbook: object, isCsv: boolean }} file
 * @param {boolean} allSheets  false uses only the first sheet
 */
export function partsOf(file, allSheets) {
  const cached = partsCache.get(file) ?? {};
  const key = allSheets ? 'all' : 'first';
  if (!cached[key]) {
    const sheetNames = allSheets ? file.workbook.SheetNames : file.workbook.SheetNames.slice(0, 1);
    cached[key] = sheetNames.map((name) => readPart(file, name)).filter(Boolean);
    partsCache.set(file, cached);
  }
  return cached[key];
}

/**
 * Merges files into one table, matching columns by name.
 * @returns {{
 *   parts: object[], allSheets: boolean, leadCount: number, headers: string[], rows: Array[],
 *   columns: { name: string, parts: Set<object> }[]   which parts each column was found in
 * }}
 */
export function mergeFiles(files, allSheets) {
  const parts = files.flatMap((file) => partsOf(file, allSheets));
  if (!parts.length) throw new Error('No data found in these files.');

  const columns = [];
  const indexByKey = new Map();
  const columnIndexes = parts.map((part) => {
    const timesSeen = new Map();
    return part.headers.map((header) => {
      let key = columnKey(header) || header.toLowerCase();
      let name = header;
      const repeat = timesSeen.get(key) ?? 0;
      timesSeen.set(key, repeat + 1);
      if (repeat) {
        // The same header twice in one sheet stays as two columns.
        key += `#${repeat + 1}`;
        name = `${header} (${repeat + 1})`;
      }
      if (!indexByKey.has(key)) {
        indexByKey.set(key, columns.length);
        columns.push({ name, parts: new Set() });
      }
      const index = indexByKey.get(key);
      columns[index].parts.add(part);
      return index;
    });
  });

  const lead = allSheets ? ['Source File', 'Source Sheet'] : ['Source File'];
  const rows = [];
  parts.forEach((part, p) => {
    for (const cells of part.rows) {
      const row = new Array(columns.length).fill('');
      columnIndexes[p].forEach((index, c) => {
        if (cells[c] !== undefined) row[index] = cells[c];
      });
      rows.push(allSheets ? [part.file, part.sheet, ...row] : [part.file, ...row]);
    }
  });

  return { parts, columns, allSheets, leadCount: lead.length, headers: [...lead, ...columns.map((c) => c.name)], rows };
}

/** Indexes of the columns that hold numbers, so they can be totalled. */
export function numericColumns({ headers, rows, leadCount }) {
  return headers.map((_, i) => i).filter((i) => {
    if (i < leadCount) return false;
    let filled = 0;
    let numbers = 0;
    for (const row of rows) {
      if (isBlank(row[i])) continue;
      filled++;
      if (numberIn(row[i]) !== undefined) numbers++;
    }
    return filled > 0 && numbers / filled >= NUMERIC_SHARE;
  });
}

/**
 * Rows and totals per group, biggest first. Group names ignore case and extra spaces.
 * @param {number} sumIndex  column to total, or -1 to count rows only
 */
export function summarise({ rows }, groupIndex, sumIndex) {
  const groups = new Map();
  for (const row of rows) {
    const label = show(row[groupIndex]).replace(/\s+/g, ' ').trim() || '(blank)';
    const key = label.toLowerCase();
    const group = groups.get(key) ?? { label, rows: 0, total: 0 };
    group.rows++;
    if (sumIndex >= 0) group.total += numberIn(row[sumIndex]) || 0;
    groups.set(key, group);
  }
  return [...groups.values()]
    .map((group) => ({ ...group, total: round2(group.total) }))
    .sort((a, b) => (sumIndex >= 0 ? b.total - a.total : b.rows - a.rows));
}

/** The summary as table rows: a header row, one row per group and a grand total. */
export function summaryTable(merged, groupIndex, sumIndex) {
  const groups = summarise(merged, groupIndex, sumIndex);
  const hasTotal = sumIndex >= 0;
  const grandTotal = round2(groups.reduce((sum, group) => sum + group.total, 0));
  return {
    head: [merged.headers[groupIndex], 'Rows', ...(hasTotal ? [`Total ${merged.headers[sumIndex]}`] : [])],
    rows: groups.map((group) => [group.label, group.rows, ...(hasTotal ? [group.total] : [])]),
    total: ['Grand total', merged.rows.length, ...(hasTotal ? [grandTotal] : [])],
  };
}
