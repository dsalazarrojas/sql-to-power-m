import { compileSqlToM, suggestFanoutRewrite } from './compiler/index.js';
import { toSqlName, workbookToHeaderContext } from './excel-schema.js';

const SAMPLE_CONTEXT = {
  tables: {
    sales: {
      excelTableName: 'tblVentas',
      columns: [
        { name: 'id', header: 'Id', internalType: 'integer' },
        { name: 'customer_id', header: 'Cliente Id', internalType: 'integer' },
        { name: 'amount', header: 'Monto Total', internalType: 'float' },
        { name: 'created_date', header: 'Fecha de Creación', internalType: 'date' },
        { name: 'status', header: 'Estado', internalType: 'string' },
      ],
      lookups: { customer_id: { table: 'customers', labelColumn: 'name', headerName: 'Customer (lookup)' } },
    },
    customers: {
      excelTableName: 'tblClientes',
      columns: [
        { name: 'id', header: 'Cliente Id', internalType: 'integer' },
        { name: 'name', header: 'Nombre Cliente', internalType: 'string' },
        { name: 'region', header: 'Región Comercial', internalType: 'string' },
      ],
    },
  },
};

const EXAMPLES = {
  'sales-summary': "SELECT status, SUM(amount) AS total, COUNT(*) AS count_rows\nFROM sales\nWHERE status IN ('paid', 'open')\nGROUP BY status\nORDER BY total DESC",
  'joined-label': 'SELECT s.customer_id, c.name AS customer_name\nFROM sales s\nLEFT JOIN customers c ON s.customer_id = c.id',
  'date-bucket': "SELECT strftime('%Y-%m', created_date) AS month, SUM(amount) AS total\nFROM sales\nGROUP BY month\nORDER BY month",
  unsupported: "WITH recent AS (SELECT * FROM sales)\nSELECT * FROM recent",
};

const state = {
  headerContext: clone(SAMPLE_CONTEXT),
  sql: EXAMPLES['sales-summary'],
  result: null,
  selectedBuilderTable: 'sales',
  schemaNote: 'Sample schema loaded. Replace it with your workbook headers before using the generated M.',
  schemaNoteTone: '',
};

const $ = selector => document.querySelector(selector);
const $$ = selector => [...document.querySelectorAll(selector)];

function clone(value) { return JSON.parse(JSON.stringify(value)); }
function encoded(value) { return encodeURIComponent(String(value)); }
function decoded(value) { return decodeURIComponent(value); }
function escapeHtml(value) { return String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char])); }
function quoteIdentifier(value) { return /^[A-Za-z_][A-Za-z0-9_$]*$/.test(String(value)) ? String(value) : `"${String(value).replaceAll('"', '""')}"`; }
function quoteSqlValue(value, column) {
  const type = String(column?.internalType || '').toLowerCase();
  if ((type.includes('int') || type.includes('float') || type.includes('number') || type.includes('decimal')) && /^[-+]?\d+(?:\.\d+)?$/.test(String(value).trim())) return String(value).trim();
  return `'${String(value).replaceAll("'", "''")}'`;
}

function reportForSql(sql) {
  const star = String(sql).match(/^\s*SELECT\s+(?:DISTINCT\s+)?\*\s+FROM\s+("(?:""|[^"])+"|[A-Za-z_][A-Za-z0-9_$]*)\b/i);
  if (!star) return { columns: [] };
  const tableName = star[1].replace(/^"|"$/g, '').replaceAll('""', '"');
  const entry = tableEntries().find(([key]) => key.toLowerCase() === tableName.toLowerCase());
  return { columns: entry ? (entry[1].columns || []).map(column => ({ name: column.name || column.header })) : [] };
}

function contextJson() { return JSON.stringify(state.headerContext, null, 2); }
function tableEntries() { return Object.entries(state.headerContext?.tables || {}); }
function getTable(key) { return state.headerContext?.tables?.[key]; }

function renderSchema() {
  const entries = tableEntries();
  $('#table-count').textContent = `${entries.length} table${entries.length === 1 ? '' : 's'}`;
  $('#schema-status').textContent = state.schemaNote;
  $('#schema-status').className = `inline-note ${state.schemaNoteTone}`.trim();
  $('#schema-json').value = contextJson();
  $('#schema-list').innerHTML = entries.length ? entries.map(([key, table]) => `
    <article class="schema-table">
      <div class="schema-table-header">
        <div class="schema-table-title"><strong title="${escapeHtml(key)}">${escapeHtml(key)}</strong><small>Excel Table: ${escapeHtml(table.excelTableName || key)}</small></div>
        <button class="table-use" type="button" data-use-table="${encoded(key)}">Use</button>
      </div>
      <div class="schema-edit-grid">
        <label class="mini-label">SQL name<input class="mini-input" data-table-field="name" data-table-key="${encoded(key)}" value="${escapeHtml(key)}"></label>
        <label class="mini-label">Excel Table name<input class="mini-input" data-table-field="excelTableName" data-table-key="${encoded(key)}" value="${escapeHtml(table.excelTableName || key)}"></label>
      </div>
      <div class="schema-fields">${(table.columns || []).map(column => `<div class="schema-field"><span title="${escapeHtml(column.name || column.header)}">${escapeHtml(column.name || column.header)}</span><span class="type-pill">${escapeHtml(column.internalType || 'text')}</span></div>`).join('')}</div>
    </article>`).join('') : '<div class="inline-note">No tables yet. Upload a workbook or paste schema JSON.</div>';
  renderBuilder();
}

function renderBuilder() {
  const entries = tableEntries();
  const select = $('#builder-table');
  const previous = state.selectedBuilderTable;
  select.innerHTML = entries.map(([key]) => `<option value="${escapeHtml(key)}">${escapeHtml(key)}</option>`).join('');
  state.selectedBuilderTable = getTable(previous) ? previous : entries[0]?.[0] || '';
  select.value = state.selectedBuilderTable;
  const table = getTable(state.selectedBuilderTable);
  const columns = table?.columns || [];
  $('#builder-columns').innerHTML = columns.map((column, index) => `<label class="column-check"><input type="checkbox" data-builder-column="${index}" checked><span>${escapeHtml(column.name || column.header)}</span></label>`).join('') || '<span class="help-text">No columns detected</span>';
  $('#builder-filter-column').innerHTML = `<option value="">No filter</option>${columns.map((column, index) => `<option value="${index}">${escapeHtml(column.name || column.header)}</option>`).join('')}`;
}

const NOT_APPLICABLE_REASON = /not the join fan-out shape|was not rejected for join fan-out/;

function suggestionBlockHtml(suggestion, { successLabel } = {}) {
  if (suggestion.ok) {
    return `<div class="suggestion-block">
        <div class="suggestion-header"><strong>${escapeHtml(successLabel || 'Suggested rewrite')}</strong><button class="secondary-button" type="button" data-copy-suggestion>Copy</button></div>
        <p class="rejection-hint">Auto-generated reshape (pre-aggregate each joined branch, then join and wrap in an outer query). Review it, then paste it into the editor above and recompile to confirm.</p>
        <pre class="m-output suggestion-output">${escapeHtml(suggestion.sql)}</pre>
      </div>`;
  }
  if (NOT_APPLICABLE_REASON.test(suggestion.reason)) {
    return `<div class="info-banner"><span class="banner-icon">i</span><span><strong>No join fan-out risk detected.</strong> This query doesn't aggregate columns from two or more joined child tables, so there's nothing to simplify.</span></div>`;
  }
  return `<div class="info-banner warning"><span class="banner-icon">!</span><span><strong>Can't safely simplify this automatically.</strong> ${escapeHtml(suggestion.reason)}</span></div>`;
}

function renderSimplifyResult(suggestion) {
  $('#result-content').innerHTML = suggestionBlockHtml(suggestion, { successLabel: 'Simplified query' });
}

function renderResult() {
  const result = state.result;
  const actions = $('#result-actions');
  if (!result) { actions.hidden = true; $('#result-content').innerHTML = '<div class="empty-result"><div class="empty-icon">⌁</div><div><strong>Ready when you are.</strong><p>Upload a workbook or use the sample schema, then compile the SQL above.</p></div></div>'; return; }
  if (result.simplifyOnly) { actions.hidden = true; renderSimplifyResult(result.suggestion); return; }
  actions.hidden = !result.ok;
  if (result.ok) {
    $('#result-content').innerHTML = `<div class="success-banner"><span class="banner-icon">✓</span><span><strong>Compiled successfully.</strong> This query reads the mapped Excel Tables in Power Query. Review the table names above before pasting.</span></div><pre class="m-output" id="m-output"></pre>`;
    $('#m-output').textContent = result.mCode;
    return;
  }
  const rejections = result.rejections || [];
  const suggestion = result.suggestion;
  const suggestionHtml = suggestion ? suggestionBlockHtml(suggestion) : '';
  $('#result-content').innerHTML = `<div class="error-banner"><span class="banner-icon">!</span><span><strong>Power Query M was not emitted.</strong> The compiler stopped safely and left your SQL unchanged.</span></div><div class="rejection-list">${rejections.map(rejection => `<article class="rejection"><div class="rejection-title"><span>${escapeHtml(rejection.construct || 'Compiler rejection')}</span><span class="position-label">source position ${Number.isFinite(rejection.position) ? rejection.position : 0}</span></div><p class="rejection-message">${escapeHtml(rejection.message || 'The SQL could not be compiled.')}</p>${rejection.hint ? `<p class="rejection-hint">Hint: ${escapeHtml(rejection.hint)}</p>` : ''}${rejection.construct === 'join fan-out' ? suggestionHtml : ''}</article>`).join('')}</div>`;
}

function compile() {
  const sql = $('#sql-input').value.trim();
  state.sql = $('#sql-input').value;
  if (!sql) { showToast('Write a SQL query first.'); $('#sql-input').focus(); return; }
  try {
    const result = compileSqlToM({ sql, report: reportForSql(sql), headerContext: state.headerContext });
    if (!result.ok && result.rejections.some(rejection => rejection.construct === 'join fan-out')) {
      result.suggestion = suggestFanoutRewrite(sql);
    }
    state.result = result;
    renderResult();
    $('#result-panel').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    state.result = { ok: false, rejections: [{ construct: 'Unexpected compiler error', position: 0, message: error.message, hint: 'This looks like a compiler issue. Keep the SQL and schema JSON if you report it.' }] };
    renderResult();
  }
}

function simplify() {
  const sql = $('#sql-input').value.trim();
  state.sql = $('#sql-input').value;
  if (!sql) { showToast('Write a SQL query first.'); $('#sql-input').focus(); return; }
  try {
    state.result = { simplifyOnly: true, suggestion: suggestFanoutRewrite(sql) };
  } catch (error) {
    state.result = { simplifyOnly: true, suggestion: { ok: false, reason: `Unexpected error: ${error.message}` } };
  }
  renderResult();
  $('#result-panel').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function loadExample() {
  const sql = EXAMPLES[$('#example-select').value] || EXAMPLES['sales-summary'];
  $('#sql-input').value = sql;
  state.result = null;
  renderResult();
  $('#sql-input').focus();
}

function insertAtCursor(text) {
  const input = $('#sql-input');
  const start = input.selectionStart ?? input.value.length;
  const end = input.selectionEnd ?? start;
  input.value = `${input.value.slice(0, start)}${text}${input.value.slice(end)}`;
  input.focus();
  input.selectionStart = input.selectionEnd = start + text.length;
}

function buildQuery() {
  const table = getTable(state.selectedBuilderTable);
  if (!table) { showToast('Load a schema before using the query starter.'); return; }
  const checked = $$('#builder-columns input:checked').map(input => table.columns[Number(input.dataset.builderColumn)]).filter(Boolean);
  const columns = checked.length ? checked : table.columns;
  const distinct = $('#builder-distinct').checked ? 'DISTINCT ' : '';
  let sql = `SELECT ${distinct}${columns.map(column => quoteIdentifier(column.name || column.header)).join(', ')}\nFROM ${quoteIdentifier(state.selectedBuilderTable)}`;
  const filterIndex = $('#builder-filter-column').value;
  const value = $('#builder-filter-value').value;
  if (filterIndex !== '' && value.trim()) {
    const column = table.columns[Number(filterIndex)];
    sql += `\nWHERE ${quoteIdentifier(column.name || column.header)} ${$('#builder-filter-op').value} ${quoteSqlValue(value, column)}`;
  }
  $('#sql-input').value = sql;
  state.result = null;
  renderResult();
  showToast('Starter query placed in the editor.');
}

function applySchemaJson() {
  try {
    const parsed = JSON.parse($('#schema-json').value);
    const tables = parsed?.tables || parsed;
    if (!tables || Array.isArray(tables) || typeof tables !== 'object' || !Object.keys(tables).length) throw new Error('Expected an object with a tables property.');
    for (const [key, table] of Object.entries(tables)) {
      if (!Array.isArray(table.columns)) throw new Error(`Table “${key}” needs a columns array.`);
    }
    state.headerContext = { tables };
    state.schemaNote = 'Schema JSON applied. Confirm each Excel Table name matches Excel before compiling.';
    state.schemaNoteTone = 'warning';
    state.result = null;
    renderSchema(); renderResult(); showToast('Schema applied.');
  } catch (error) {
    showToast(`Schema JSON error: ${error.message}`);
  }
}

async function readWorkbook(file) {
  if (!window.XLSX) { showToast('The workbook reader is still loading. Try again in a moment, or use schema JSON.'); return; }
  if (file.size > 30 * 1024 * 1024) { showToast('For privacy and browser performance, keep workbooks below 30 MB.'); return; }
  try {
    const data = await file.arrayBuffer();
    const workbook = window.XLSX.read(data, { type: 'array', cellDates: true, bookFiles: true });
    const extracted = workbookToHeaderContext(workbook, window.XLSX);
    state.headerContext = extracted.headerContext;
    state.schemaNote = `${file.name} read locally. ${extracted.warnings.join(' ') || 'Review the detected mappings before compiling.'}`;
    state.schemaNoteTone = extracted.warnings.length ? 'warning' : '';
    state.result = null;
    renderSchema(); renderResult();
    showToast(`${Object.keys(state.headerContext.tables).length} table(s) detected.`);
  } catch (error) {
    state.schemaNote = `Could not read ${file.name}: ${error.message}`;
    state.schemaNoteTone = 'error';
    renderSchema();
    showToast('Workbook could not be read. Try schema JSON instead.');
  }
}

function loadSample() {
  state.headerContext = clone(SAMPLE_CONTEXT);
  state.schemaNote = 'Sample schema loaded. Replace it with your workbook headers before using the generated M.';
  state.schemaNoteTone = '';
  state.result = null;
  renderSchema(); renderResult(); showToast('Sample schema loaded.');
}

function updateTableField(input) {
  const oldKey = decoded(input.dataset.tableKey);
  const field = input.dataset.tableField;
  const table = getTable(oldKey);
  if (!table) return;
  const value = input.value.trim();
  if (!value) { renderSchema(); return; }
  if (field === 'excelTableName') table.excelTableName = value;
  if (field === 'name') {
    const newKey = toSqlName(value, oldKey);
    if (newKey !== oldKey && getTable(newKey)) { showToast(`SQL table “${newKey}” already exists.`); renderSchema(); return; }
    const tables = {};
    for (const [key, item] of tableEntries()) tables[key === oldKey ? newKey : key] = item;
    state.headerContext.tables = tables;
    state.selectedBuilderTable = newKey;
  }
  state.result = null;
  renderSchema(); renderResult();
}

function showToast(message) {
  const toast = $('#toast');
  toast.textContent = message;
  toast.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove('show'), 3300);
}

async function copyM() {
  if (!state.result?.mCode) return;
  try { await navigator.clipboard.writeText(state.result.mCode); showToast('Power Query M copied to the clipboard.'); }
  catch { showToast('Clipboard access failed. Select the M output and copy it manually.'); }
}

function downloadM() {
  if (!state.result?.mCode) return;
  const url = URL.createObjectURL(new Blob([state.result.mCode], { type: 'text/plain;charset=utf-8' }));
  const anchor = document.createElement('a'); anchor.href = url; anchor.download = 'query.pq'; anchor.click(); URL.revokeObjectURL(url);
  showToast('query.pq downloaded.');
}

function wireEvents() {
  $('#sql-input').value = state.sql;
  $('#choose-file').addEventListener('click', () => $('#workbook-file').click());
  $('#workbook-file').addEventListener('change', event => { const file = event.target.files?.[0]; if (file) readWorkbook(file); event.target.value = ''; });
  $('#load-sample').addEventListener('click', loadSample);
  $('#apply-schema').addEventListener('click', applySchemaJson);
  $('#compile-button').addEventListener('click', compile);
  $('#simplify-button').addEventListener('click', simplify);
  $('#load-example').addEventListener('click', loadExample);
  $('#build-query').addEventListener('click', buildQuery);
  $('#copy-m').addEventListener('click', copyM);
  $('#download-m').addEventListener('click', downloadM);
  $('#builder-table').addEventListener('change', event => { state.selectedBuilderTable = event.target.value; renderBuilder(); });
  $('#schema-list').addEventListener('click', event => { const button = event.target.closest('[data-use-table]'); if (button) insertAtCursor(`\nFROM ${quoteIdentifier(decoded(button.dataset.useTable))}`); });
  $('#schema-list').addEventListener('change', event => { if (event.target.matches('[data-table-field]')) updateTableField(event.target); });
  $('#sql-input').addEventListener('keydown', event => { if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') { event.preventDefault(); compile(); } });
  $('#result-content').addEventListener('click', async event => {
    if (!event.target.closest('[data-copy-suggestion]')) return;
    const sql = state.result?.suggestion?.sql;
    if (!sql) return;
    try { await navigator.clipboard.writeText(sql); showToast('Suggested rewrite copied to the clipboard.'); }
    catch { showToast('Clipboard access failed. Select the suggested SQL and copy it manually.'); }
  });
  const dropzone = $('#dropzone');
  for (const type of ['dragenter', 'dragover']) dropzone.addEventListener(type, event => { event.preventDefault(); dropzone.classList.add('is-dragging'); });
  for (const type of ['dragleave', 'drop']) dropzone.addEventListener(type, event => { event.preventDefault(); dropzone.classList.remove('is-dragging'); });
  dropzone.addEventListener('drop', event => { const file = event.dataTransfer.files?.[0]; if (file) readWorkbook(file); });
}

renderSchema();
renderResult();
wireEvents();
