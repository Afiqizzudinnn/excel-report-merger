/* global XLSX */
/**
 * Page code for the Excel Report Merger: collect files, merge them, show the result and download it.
 * The merging rules themselves live in merge.js.
 */
import { downloadWorkbook, readWorkbook, show } from './excel.js';
import { columnKey, mergeFiles, numericColumns, partsOf, summaryTable } from './merge.js';
import { sampleFiles } from './sample.js';

const PREVIEW_ROW_LIMIT = 200;
// Columns picked by default for the summary, when the files have them.
const GROUP_BY_HINT = /branch|region|outlet|store|department|dept|team|category/;
const TOTAL_OF_HINT = /sales|amount|total|revenue|value|price/;
const FILE_ICON = '<span class="file-icon"><svg viewBox="0 0 20 20" aria-hidden="true"><rect x="3.5" y="3.5" width="13" height="13" rx="2"/><path d="M3.5 8h13M3.5 12h13M8 8v8.5"/></svg></span>';

const $ = (selector) => document.querySelector(selector);

const el = {
  pick: $('#pick'),
  fileInput: $('#file-input'),
  dropZone: $('#drop-zone'),
  sample: $('#sample'),
  fileCount: $('#file-count'),
  fileList: $('#file-list'),
  allSheets: $('#opt-all-sheets'),
  merge: $('#merge'),
  notice: $('#notice'),
  results: $('#results'),
  summaryLine: $('#summary-line'),
  download: $('#download'),
  gapWarning: $('#gap-warning'),
  tabs: $('#tabs'),
  pickers: $('#pickers'),
  groupBy: $('#group-by'),
  totalOf: $('#total-of'),
  panes: { summary: $('#summary'), columns: $('#column-check'), rows: $('#preview') },
  footnote: $('#footnote'),
};

const state = { files: [], merged: null };

const escapeHtml = (text) => String(text).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
const plural = (count, word) => `<b>${count}</b> ${word}${count === 1 ? '' : 's'}`;
const formatNumber = (n) => n.toLocaleString('en-MY', { maximumFractionDigits: 2 });
const partName = (part) => (state.merged.allSheets ? `${part.file} / ${part.sheet}` : part.file);
const options = (indexes, headers) => indexes.map((i) => `<option value="${i}">${escapeHtml(headers[i])}</option>`).join('');

function showNotice(message) {
  el.notice.textContent = message;
  el.notice.hidden = !message;
}

// ---------- Files ----------

function renderFiles() {
  const count = state.files.length;
  el.fileList.innerHTML = count
    ? state.files.map((file, i) => {
      const rows = partsOf(file, el.allSheets.checked).reduce((sum, part) => sum + part.rows.length, 0);
      const sheets = file.workbook.SheetNames.length;
      return `<li>
        ${FILE_ICON}<span class="file-name">${escapeHtml(file.name)}</span>
        <span class="file-meta">${rows} rows${sheets > 1 ? ` · ${sheets} sheets` : ''}</span>
        <button class="file-remove" type="button" data-index="${i}" aria-label="Remove ${escapeHtml(file.name)}">×</button>
      </li>`;
    }).join('')
    : '<li class="is-empty">Your files will show up here</li>';
  el.fileCount.textContent = count ? `${count} ${count === 1 ? 'file' : 'files'} ready` : 'No files yet';
  el.fileCount.classList.toggle('is-loaded', count > 0);
  el.merge.disabled = count === 0;
  el.results.hidden = true;
}

async function addFiles(fileList) {
  showNotice('');
  for (const file of fileList) {
    try {
      const { workbook, isCsv } = await readWorkbook(file);
      state.files.push({ name: file.name, workbook, isCsv });
    } catch {
      showNotice(`Sorry, "${file.name}" could not be opened. Please use .xlsx, .xls or .csv files.`);
    }
  }
  renderFiles();
}

function loadSample({ scroll }) {
  state.files = sampleFiles();
  renderFiles();
  runMerge({ scroll });
}

// ---------- Merging and rendering ----------

function runMerge({ scroll }) {
  try {
    state.merged = mergeFiles(state.files, el.allSheets.checked);
  } catch (error) {
    showNotice(error.message);
    return;
  }
  const merged = state.merged;
  el.summaryLine.innerHTML = `${plural(merged.parts.length, merged.allSheets ? 'sheet' : 'file')}, ${plural(merged.rows.length, 'row')} and ${plural(merged.columns.length, 'column')}, lined up by column name.`;

  setUpPickers(merged);
  renderSummary();
  renderColumnCheck(merged);
  renderPreview(merged);

  const notes = [];
  if (merged.rows.length > PREVIEW_ROW_LIMIT) notes.push(`Showing the first ${PREVIEW_ROW_LIMIT} of ${merged.rows.length} rows.`);
  notes.push('The download has 4 sheets: Merged, Summary, Column Check and Files.');
  el.footnote.textContent = notes.join(' ');

  el.results.hidden = false;
  if (scroll) el.results.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function setUpPickers(merged) {
  const numeric = numericColumns(merged);
  const textual = merged.headers.map((_, i) => i).filter((i) => !numeric.includes(i));
  const pick = (indexes, hint) => indexes.find((i) => hint.test(columnKey(merged.headers[i]))) ?? indexes[0];

  el.groupBy.innerHTML = options(textual, merged.headers);
  el.groupBy.value = pick(textual, GROUP_BY_HINT);
  el.totalOf.innerHTML = `<option value="-1">(count rows only)</option>${options(numeric, merged.headers)}`;
  el.totalOf.value = numeric.length ? pick(numeric, TOTAL_OF_HINT) : -1;
}

const currentSummary = () => summaryTable(state.merged, Number(el.groupBy.value), Number(el.totalOf.value));

function renderSummary() {
  const { head, rows, total } = currentSummary();
  const cell = (value) => (typeof value === 'number' ? `<td class="num">${formatNumber(value)}</td>` : `<td>${escapeHtml(value)}</td>`);
  el.panes.summary.innerHTML = `<table class="data-table">
    <thead><tr>${head.map((h, i) => `<th${i ? ' class="num"' : ''}>${escapeHtml(h)}</th>`).join('')}</tr></thead>
    <tbody>
      ${rows.map((row) => `<tr>${row.map(cell).join('')}</tr>`).join('')}
      <tr class="total">${total.map(cell).join('')}</tr>
    </tbody>
  </table>`;
}

function renderColumnCheck(merged) {
  const head = merged.parts.map((part) => `<th class="center">${escapeHtml(partName(part))}</th>`).join('');
  const body = merged.columns.map((column) => {
    const cells = merged.parts.map((part) => (column.parts.has(part) ? '<td class="is-present">✓</td>' : '<td class="is-missing">missing</td>'));
    return `<tr><td>${escapeHtml(column.name)}</td>${cells.join('')}</tr>`;
  });
  el.panes.columns.innerHTML = `<table class="data-table"><thead><tr><th>Column</th>${head}</tr></thead><tbody>${body.join('')}</tbody></table>`;

  const unit = merged.allSheets ? 'sheets' : 'files';
  const gaps = merged.columns.filter((column) => column.parts.size < merged.parts.length);
  el.gapWarning.innerHTML = gaps.map((column) => `<b>${escapeHtml(column.name)}</b> is only in ${column.parts.size} of ${merged.parts.length} ${unit}, so the other rows are blank there. If it is the same as another column with a different name, rename it in the source file and merge again.`).join('<br>');
  el.gapWarning.hidden = gaps.length === 0;
}

function renderPreview(merged) {
  const head = merged.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('');
  const body = merged.rows.slice(0, PREVIEW_ROW_LIMIT)
    .map((row) => `<tr>${row.map((value) => `<td${typeof value === 'number' ? ' class="num"' : ''}>${escapeHtml(show(value))}</td>`).join('')}</tr>`);
  el.panes.rows.innerHTML = `<table class="data-table"><thead><tr>${head}</tr></thead><tbody>${body.join('')}</tbody></table>`;
}

function showTab(tab) {
  Object.entries(el.panes).forEach(([name, pane]) => { pane.hidden = name !== tab; });
  el.pickers.hidden = tab !== 'summary';
}

// ---------- Download ----------

function downloadResult() {
  const merged = state.merged;
  const summary = currentSummary();
  downloadWorkbook('merged_report.xlsx', [
    ['Merged', [merged.headers, ...merged.rows]],
    ['Summary', [summary.head, ...summary.rows, summary.total]],
    ['Column Check', [
      ['Column', ...merged.parts.map(partName)],
      ...merged.columns.map((column) => [column.name, ...merged.parts.map((part) => (column.parts.has(part) ? 'Yes' : 'Missing'))]),
    ]],
    ['Files', [
      ['File', 'Sheet', 'Rows'],
      ...merged.parts.map((part) => [part.file, part.sheet, part.rows.length]),
      ['Merged on', new Date().toLocaleString('en-GB'), ''],
    ]],
  ]);
}

// ---------- Start ----------

function init() {
  el.pick.addEventListener('click', () => el.fileInput.click());
  el.fileInput.addEventListener('change', () => {
    addFiles([...el.fileInput.files]);
    el.fileInput.value = '';
  });

  ['dragenter', 'dragover'].forEach((type) => el.dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    el.dropZone.classList.add('is-dragover');
  }));
  ['dragleave', 'drop'].forEach((type) => el.dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    el.dropZone.classList.remove('is-dragover');
  }));
  el.dropZone.addEventListener('drop', (event) => addFiles([...event.dataTransfer.files]));

  el.fileList.addEventListener('click', (event) => {
    const button = event.target.closest('[data-index]');
    if (!button) return;
    state.files.splice(Number(button.dataset.index), 1);
    renderFiles();
  });
  el.allSheets.addEventListener('change', renderFiles);
  el.sample.addEventListener('click', () => loadSample({ scroll: true }));
  el.merge.addEventListener('click', () => runMerge({ scroll: true }));
  el.groupBy.addEventListener('change', renderSummary);
  el.totalOf.addEventListener('change', renderSummary);
  el.download.addEventListener('click', downloadResult);
  el.tabs.addEventListener('click', (event) => {
    const button = event.target.closest('button');
    if (!button) return;
    el.tabs.querySelectorAll('button').forEach((b) => b.setAttribute('aria-selected', String(b === button)));
    showTab(button.dataset.value);
  });

  renderFiles();
  if (typeof XLSX === 'undefined') {
    showNotice('The spreadsheet library did not load. Check your internet connection and refresh the page.');
    return;
  }
  if (new URLSearchParams(window.location.search).has('demo')) loadSample({ scroll: false });
}

init();
