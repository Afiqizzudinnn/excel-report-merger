/* global XLSX */
/**
 * Three made-up branch sales reports for the "Try it with 3 sample branch reports" button.
 * Penang uses a different column order and casing, and Johor has an extra column,
 * so the sample shows the column matching and the column check at work.
 */
import { sheetFromRows, ymdToSerial } from './excel.js';

const PRICES = { 'Laptop Stand': 89, 'Wireless Mouse': 45, 'USB-C Hub': 129, 'Office Chair': 459, 'Monitor 24"': 599 };

const marchDate = (day) => ({ date: true, serial: ymdToSerial(2026, 3, day), z: 'dd/mm/yyyy' });

// [day of March, product, quantity, salesperson]
const KUALA_LUMPUR = [[2, 'Office Chair', 4], [3, 'Wireless Mouse', 20], [5, 'Monitor 24"', 6], [9, 'USB-C Hub', 12], [12, 'Laptop Stand', 15], [16, 'Office Chair', 3]];
const PENANG = [[2, 'Wireless Mouse', 14], [4, 'Laptop Stand', 9], [10, 'Monitor 24"', 4], [11, 'USB-C Hub', 7], [18, 'Office Chair', 2]];
const JOHOR = [[3, 'Monitor 24"', 5, 'Farah'], [6, 'Wireless Mouse', 18, 'Kumar'], [9, 'Office Chair', 5, 'Farah'], [13, 'Laptop Stand', 11, 'Jason'], [17, 'USB-C Hub', 9, 'Kumar']];

function sampleFile(name, rows) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheetFromRows(rows), 'March');
  // Round-trip through .xlsx so the sample behaves exactly like an uploaded file.
  const bytes = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
  return { name, workbook: XLSX.read(bytes, { type: 'array', cellNF: true }), isCsv: false };
}

export function sampleFiles() {
  return [
    sampleFile('Sales_KL_Mar2026.xlsx', [
      ['Date', 'Branch', 'Product', 'Qty', 'Sales (RM)'],
      ...KUALA_LUMPUR.map(([day, product, qty]) => [marchDate(day), 'Kuala Lumpur', product, qty, qty * PRICES[product]]),
    ]),
    sampleFile('Sales_Penang_Mar2026.xlsx', [
      ['product', 'date', 'branch', 'qty', 'SALES (RM)'],
      ...PENANG.map(([day, product, qty]) => [product, marchDate(day), 'Penang', qty, qty * PRICES[product]]),
    ]),
    sampleFile('Sales_Johor_Mar2026.xlsx', [
      ['Date', 'Branch', 'Product', 'Qty', 'Sales (RM)', 'Salesperson'],
      ...JOHOR.map(([day, product, qty, person]) => [marchDate(day), 'Johor Bahru', product, qty, qty * PRICES[product], person]),
    ]),
  ];
}
