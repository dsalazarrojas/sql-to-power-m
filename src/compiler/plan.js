'use strict';
const { containsAggregate } = require('./validate');
function makePlan(ast, headerContext, report) {
  if (ast.type === 'union') return { kind: 'union', ast, arms: ast.arms.map(arm => makePlan(arm, headerContext, { columns: arm.select.map(x => ({ name: x.alias })) })) };
  const steps = []; const base = ast.from;
  // A collapsed lookup is stored as a materialized column on the base worksheet,
  // rather than as a column on the parent table.  Do this before planning any
  // expression-consuming step so filters, groups, and aggregates see that header.
  const collapsedLookups = new Map();
  for (const join of ast.joins) {
    const lookup = collapse(join, ast, headerContext);
    if (lookup) collapsedLookups.set(join, lookup);
  }
  for (const [join, lookup] of collapsedLookups) substituteCollapsedLookup(ast, join, lookup);
  steps.push({ name: 'Source', kind: 'source', table: base.resolvedExcelTable || base.name, subplan: base.subquery && makePlan(base.subquery, headerContext, { columns: base.subquery.select.map(x => ({ name: x.alias })) }) });
  const types = collectTypes(ast, headerContext); if (types.length) steps.push({ name: 'ChangedType', kind: 'types', types });
  const computed = ast.select.filter(x => (x.expression.type !== 'column' && !containsAggregate(x.expression)) || (x.expression.type === 'call' && x.expression.name === 'SUM' && x.expression.args[0] && x.expression.args[0].type === 'case'));
  const baseColumns = ast.from.subquery ? ast.from.subquery.select.map(x => ({ name: x.alias, header: x.alias })) : ((((headerContext || {}).tables || {})[ast.from.name] || {}).columns || []);
  const occupied = baseColumns.concat(...ast.joins.map(j => j.table.subquery ? j.table.subquery.select.map(x => ({ name: x.alias, header: x.alias })) : (((headerContext || {}).tables || {})[j.table.name] || {}).columns || []));
  computed.forEach((item) => { item.outputName = occupied.some(c => [typeof c === 'string' ? c : (c.name || c.originalName), typeof c === 'string' ? c : (c.header || c.displayHeader || c.name || c.originalName)].some(n => String(n).toLowerCase() === String(item.alias).toLowerCase())) ? `${item.alias} Reporte` : item.alias; });
  const joinEquality = (on, alias) => { if (!on || typeof on !== 'object') return null; if (on.type === 'binary' && on.op === 'AND') return joinEquality(on.left, alias) || joinEquality(on.right, alias); return on.type === 'binary' && on.op === '=' && on.left && on.left.type === 'column' && on.right && on.right.type === 'column' && [on.left, on.right].some(x => String(x.resolvedAlias || x.table).toLowerCase() === String(alias).toLowerCase()) ? on : null; }; const withoutJoinEquality = (on, equality) => { if (on === equality) return null; if (on && on.type === 'binary' && on.op === 'AND') { const left = withoutJoinEquality(on.left, equality); const right = withoutJoinEquality(on.right, equality); return left && right ? { ...on, left, right } : left || right; } return on; };
  let n = 0; for (const join of ast.joins) { if (!collapsedLookups.has(join)) { n++; const equality = join.on && join.on.type === 'binary' && join.on.op === 'AND' && joinEquality(join.on, join.table.alias); const remainder = equality && withoutJoinEquality(join.on, equality); if (equality && remainder) join.on = equality; let subplan = join.table.subquery && makePlan(join.table.subquery, headerContext, { columns: join.table.subquery.select.map(x => ({ name: x.alias })) }); if (remainder && subplan) subplan.steps.push({ name: 'FilteredRows', kind: 'filter', expression: remainder }); steps.push({ name: `Joined${n}`, kind: 'join', join, subplan }); steps.push({ name: `Expanded${n}`, kind: 'expand', join }); if (remainder && !subplan) steps.push({ name: `JoinedFilter${n}`, kind: 'filter', expression: remainder }); } }
  // A predicate can reference a joined table (notably LEFT JOIN ... IS NULL),
  // so it must run after joins have made those columns available.
  if (ast.where) steps.push({ name: 'FilteredRows', kind: 'filter', expression: ast.where });
  computed.forEach((item, i) => steps.push({ name: `Added${i + 1}`, kind: 'add', item }));
  const aggregate = ast.select.some(x => containsAggregate(x.expression)); if (aggregate) { steps.push({ name: 'GroupedRows', kind: 'group', ast }); const groupRenames = ast.select.filter(x => x.expression.type === 'column').map(x => [groupOutputName(x.expression, ast), x.alias]).filter(x => x[0] !== x[1]); if (groupRenames.length) steps.push({ name: 'RenamedGroupColumns', kind: 'rename', pairs: groupRenames }); if (ast.having) steps.push({ name: 'HavingRows', kind: 'having', expression: ast.having }); }
  if (ast.distinct) steps.push({ name: 'CombinedRows', kind: 'distinct' });
  const renames = computed.filter(x => x.outputName !== x.alias).map(x => [x.outputName, x.alias]);
  // A computed alias wins over a physical column with the same SQL output name.
  // Leave that physical column under its source header; SelectedColumns drops it.
  const computedTargets = new Set(renames.map(x => String(x[1]).toLowerCase()));
  if (!aggregate) renames.push(...ast.select.filter(x => x.expression.type === 'column' && (String(x.expression.resolvedAlias || x.expression.table).toLowerCase() === String(ast.from.alias).toLowerCase() || x.expression.collapsedLookupHeader)).map(x => [x.expression.resolvedHeader || x.expression.name, x.alias]).filter(x => x[0] !== x[1] && !computedTargets.has(String(x[1]).toLowerCase())));
  if (renames.length) steps.push({ name: 'RenamedColumns', kind: 'rename', pairs: renames });
  const normalize = s => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const headerRenamePairs = []; const selectedColumns = report && Array.isArray(report.columns) && report.columns.length ? report.columns.map(x => { const column = typeof x === 'string' ? { field: x } : x || {}; const selected = ast.select.find(y => column.field != null && normalize(y.alias) === normalize(column.field)) || (typeof x !== 'string' && column.name != null && ast.select.find(y => normalize(y.alias) === normalize(column.name))); if (!selected) { const value = column.field != null ? column.field : column.name; const err = new Error('report.columns entry has no matching SQL SELECT column/alias'); err.rejection = { code: 'validation-error', construct: 'SelectedColumns', position: 0, message: `report.columns entry "${value}" has no matching SQL SELECT column/alias; cannot compile a Power Query projection for it.`, hint: 'Ensure report.columns[].field matches a SELECT column alias (case/diacritic-insensitive), or use queries.excel_power_query for raw M.' }; throw err; } const header = typeof column.label === 'string' && column.label.trim() ? column.label : column.name; if (header != null && String(header) !== String(selected.alias)) headerRenamePairs.push([selected.alias, header]); return selected.alias; }) : ast.select.map(x => x.alias);
  steps.push({ name: 'SelectedColumns', kind: 'select', columns: selectedColumns });
  if (ast.orderBy.length) steps.push({ name: 'SortedRows', kind: 'sort', orderBy: ast.orderBy }); if (headerRenamePairs.length) steps.push({ name: 'RelabeledColumns', kind: 'rename', pairs: headerRenamePairs }); return { steps, ast };
}
function collapse(join, ast, context) {
  if (join.kind === 'CROSS') return null;
  const lookup = lookupForJoin(join, ast, context);
  const references = [];
  const collect = e => { if (!e || typeof e !== 'object') return; if (e.type === 'column') { if (String(e.resolvedAlias || e.table).toLowerCase() === String(join.table.alias).toLowerCase()) references.push(e); return; } Object.values(e).forEach(v => Array.isArray(v) ? v.forEach(collect) : collect(v)); };
  ast.select.forEach(x => collect(x.expression)); collect(ast.where); (ast.groupBy || []).forEach(collect); collect(ast.having); (ast.orderBy || []).forEach(x => collect(x.expression));
  // An unused LEFT JOIN cannot alter the base row set, so it can be elided.
  // INNER and FULL joins retain existence/multiplicity semantics even with no
  // projected child columns and must remain in the pipeline.
  if (!references.length && join.kind === 'LEFT') return {};
  // Do not collapse an intermediate join: a following join may need one of its
  // columns as its key, even when SELECT only uses the eventual child table.
  const joinAlias = String(join.table.alias).toLowerCase();
  if (ast.joins.some(other => other !== join && other.on && [other.on.left, other.on.right].some(x => x && x.type === 'column' && String(x.resolvedAlias || x.table).toLowerCase() === joinAlias))) return null;
  return lookup && references.every(x => String(x.name).toLowerCase() === String(lookup.labelColumn).toLowerCase()) ? lookup : null;
}
function lookupForJoin(join, ast, context) {
  if (!join.on) return null;
  const lookups = ((((context || {}).tables || {})[ast.from.name] || {}).lookups || {});
  const baseAlias = String(ast.from.alias).toLowerCase(); const joinAlias = String(join.table.alias).toLowerCase();
  const onColumns = [join.on.left, join.on.right];
  const baseColumn = onColumns.find(x => x && x.type === 'column' && String(x.resolvedAlias || x.table).toLowerCase() === baseAlias);
  const parentColumn = onColumns.find(x => x && x.type === 'column' && String(x.resolvedAlias || x.table).toLowerCase() === joinAlias);
  if (!baseColumn || !parentColumn) return null;
  return Object.entries(lookups).map(([fkColumn, value]) => ({ ...value, fkColumn })).find(l => String(l.fkColumn).toLowerCase() === String(baseColumn.name).toLowerCase() && String(l.table).toLowerCase() === String(join.table.name).toLowerCase() && l.headerName);
}
function substituteCollapsedLookup(ast, join, lookup) {
  const visit = e => { if (!e || typeof e !== 'object') return; if (e.type === 'column') { if (String(e.resolvedAlias || e.table).toLowerCase() === String(join.table.alias).toLowerCase() && String(e.name).toLowerCase() === String(lookup.labelColumn).toLowerCase()) { e.resolvedHeader = lookup.headerName; e.collapsedLookupHeader = lookup.headerName; } return; } Object.values(e).forEach(v => Array.isArray(v) ? v.forEach(visit) : visit(v)); };
  ast.select.forEach(x => visit(x.expression)); visit(ast.where); (ast.groupBy || []).forEach(visit); visit(ast.having); (ast.orderBy || []).forEach(x => visit(x.expression));
}
function groupOutputName(e, ast) { const key = x => JSON.stringify(x, (k, v) => ['resolvedTable', 'resolvedAlias', 'resolvedHeader'].includes(k) ? undefined : v); const item = ast.select.find(x => key(x.expression) === key(e)); const joined = e.type === 'column' && !e.collapsedLookupHeader && String(e.resolvedAlias || e.table).toLowerCase() !== String(ast.from.alias).toLowerCase(); return item && (e.type !== 'column' || joined) ? (item.outputName || item.alias) : (e.resolvedHeader || e.name); }
function collectTypes(ast, context) { const out = []; const tables = (context || {}).tables || {}; const needs = e => e && (e.type === 'cast' || (e.type === 'call' && ['STRFTIME','JULIANDAY'].includes(e.name)) || (e.type === 'binary' && ['+','-','*','/'].includes(e.op))) || (e && Object.values(e).some(v => Array.isArray(v) ? v.some(needs) : v && typeof v === 'object' && needs(v))); const dateCall = e => e && e.type === 'call' && ['DATE','JULIANDAY','STRFTIME'].includes(e.name); const comparisonColumns = e => !e || typeof e !== 'object' ? [] : e.type === 'binary' && ['<','>','<=','>='].includes(e.op) ? e.left.type === 'column' && dateCall(e.right) ? [e.left] : e.right.type === 'column' && dateCall(e.left) ? [e.right] : [] : Object.values(e).flatMap(v => Array.isArray(v) ? v.flatMap(comparisonColumns) : comparisonColumns(v));
  for (const item of ast.select) if (needs(item.expression) && item.expression.type === 'column') { const c = ((tables[item.expression.resolvedTable] || {}).columns || []).find(x => (x.name || x) === item.expression.name); const type = String((c && c.internalType) || '').toLowerCase(); const header = item.expression.resolvedHeader || item.expression.name; out.push([header, /date|fecha/.test(header.toLowerCase()) ? 'date' : (/int|float|decimal|number/.test(type) ? 'number' : 'text')]); } [ast.where, ast.having].flatMap(comparisonColumns).forEach(column => out.push([column.resolvedHeader || column.name, 'date'])); return out; }
module.exports = { makePlan };
