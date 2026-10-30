/**
 * Reading text that should really be a date or a number.
 * Shared by the Data Cleaner and the Report Merger, so keep this file identical across those repos.
 */

const MONTHS = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12,
  january: 1, february: 2, march: 3, april: 4, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
  // Malay month names that differ from English
  mac: 3, mei: 5, ogos: 8, okt: 10, dis: 12,
};

const MIN_YEAR = 1900;
const MAX_YEAR = 2100;
// Plain digit strings longer than this are IDs (IC, account, card numbers), not amounts.
const MAX_PLAIN_DIGITS = 9;

/** Two-digit years: 00-49 become 2000s, 50-99 become 1900s. */
const fullYear = (text) => {
  const year = Number(text);
  if (year >= 100) return year;
  return year < 50 ? 2000 + year : 1900 + year;
};

function validDate(year, month, day) {
  if (month < 1 || month > 12 || day < 1 || year < MIN_YEAR || year > MAX_YEAR) return null;
  const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return day <= daysInMonth ? { y: year, m: month, d: day } : null;
}

/**
 * Reads a date written as text.
 * Handles 2026-03-03, 03/03/2026, 3-3-26, 3 Mar 2026, 11 Mac 2026 and Mar 9, 2026.
 * @param {string} text
 * @param {boolean} dayFirst  how to read 03/04/2026 (true: 3 April, false: 4 March)
 * @returns {{ y: number, m: number, d: number } | null}
 */
export function parseDate(text, dayFirst) {
  const s = text.trim();
  let m;

  if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/))) {
    return validDate(Number(m[1]), Number(m[2]), Number(m[3]));
  }
  if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/))) {
    const a = Number(m[1]);
    const b = Number(m[2]);
    const year = fullYear(m[3]);
    // Use the chosen order; fall back to the other only when the chosen one cannot exist (e.g. month 31).
    return dayFirst
      ? validDate(year, b, a) || validDate(year, a, b)
      : validDate(year, a, b) || validDate(year, b, a);
  }
  if ((m = s.match(/^(\d{1,2})[\s-]+([A-Za-z]{3,9})\.?,?[\s-]+(\d{2}|\d{4})$/))) {
    const month = MONTHS[m[2].toLowerCase()];
    return month ? validDate(fullYear(m[3]), month, Number(m[1])) : null;
  }
  if ((m = s.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$/))) {
    const month = MONTHS[m[1].toLowerCase()];
    return month ? validDate(Number(m[3]), month, Number(m[2])) : null;
  }
  return null;
}

/**
 * Reads a number written as text, such as "RM 1,250.00", "(120.00)" or " 89.90".
 * Phone numbers, IDs with a leading zero and long digit strings are left alone.
 * @returns {{ value: number, formatted: boolean } | null}
 *   `formatted` is true when the text had a currency, thousands commas or accounting brackets.
 */
export function parseNumber(text) {
  let t = text.trim();
  let negative = false;
  let formatted = false;

  if (/^\(.+\)$/.test(t)) {
    negative = true;
    formatted = true;
    t = t.slice(1, -1).trim();
  }
  if (t[0] === '-') {
    negative = !negative;
    t = t.slice(1).trim();
  }
  const currency = t.match(/^(RM|MYR|SGD|S\$|USD|US\$|\$)\s*/i);
  if (currency) {
    formatted = true;
    t = t.slice(currency[0].length);
  }
  if (t[0] === '-') {
    negative = !negative;
    t = t.slice(1);
  }

  if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) {
    formatted = true;
    t = t.replace(/,/g, '');
  } else if (!/^\d+(\.\d+)?$/.test(t)) {
    return null;
  }

  const wholePart = t.split('.')[0];
  if (wholePart.length > 1 && wholePart[0] === '0') return null;
  if (!formatted && wholePart.length > MAX_PLAIN_DIGITS) return null;

  const value = Number.parseFloat(t);
  return { value: negative ? -value : value, formatted };
}
