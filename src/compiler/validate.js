'use strict';
const aggregateNames = new Set(['SUM', 'COUNT', 'AVG', 'MIN', 'MAX']);
const reject = (construct, position, message, hint) => ({ ok: false, rejection: { code: 'validation-error', construct, position: position == null ? 0 : position, message: message || `${construct} is not supported`, hint: hint || 'Simplify the SQL, or check docs/excel-report-sql-patterns.md for a documented rewrite pattern. queries.excel_power_query (hand-written Power Query M) is a last resort for cases with no SQL equivalent — not a default fallback; most rejections have a documented SQL rewrite.' } });
function validate(ast, headerContext, report) {
  if (ast.type === 'union') { const checked = ast.arms.map(arm => validate(arm, headerContext, report)); const bad = checked.find(x => !x.ok); if (bad) return bad; const names = arm => arm.select.map(x => x.alias.toLowerCase()); const expected = names(ast.arms[0]); for (let i = 1; i < ast.arms.length; i++) if (expected.length !== names(ast.arms[i]).length || expected.some((x, j) => x !== names(ast.arms[i])[j])) return reject('UNION column sets', 0, 'UNION arms must project the same column aliases in the same order'); return { ok: true, ast, aliases: {} }; }
  // headerContext.tables is keyed by raw SQL table name; excelTableName and column.header are optional Excel-facing names.
  const tables = (headerContext && headerContext.tables) || {}; const find = n => Object.keys(tables).find(k => k.toLowerCase() === String(n).toLowerCase());
  const outputSpec = table => ({ columns: table.subquery.select.map(x => ({ name: x.alias, header: x.alias })) });
  if (ast.from.subquery) { const r = validate(ast.from.subquery, headerContext, { columns: ast.from.subquery.select.map(x => ({ name: x.alias })) }); if (!r.ok) return r; ast.from.resolvedExcelTable = null; }
  const baseKey = ast.from.subquery ? ast.from.alias : find(ast.from.name);
  if (!baseKey) return reject(`Unknown table ${ast.from.name}`, ast.from.position, `Unknown table ${ast.from.name}`);
  const aliases = { [ast.from.alias.toLowerCase()]: { key: baseKey, spec: ast.from.subquery ? outputSpec(ast.from) : tables[baseKey] } };
  if (!ast.from.subquery) ast.from.resolvedExcelTable = tables[baseKey].excelTableName || ast.from.name;
  for (const join of ast.joins) { if (join.table.subquery) { const r = validate(join.table.subquery, headerContext, { columns: join.table.subquery.select.map(x => ({ name: x.alias })) }); if (!r.ok) return r; aliases[join.table.alias.toLowerCase()] = { key: join.table.alias, spec: outputSpec(join.table) }; continue; } const key = find(join.table.name); if (!key) return reject(`Unknown table ${join.table.name}`, join.table.position, `Unknown table ${join.table.name}`); join.table.resolvedExcelTable = tables[key].excelTableName || join.table.name; aliases[join.table.alias.toLowerCase()] = { key, spec: tables[key] }; }
  const cols = spec => (spec.columns || []).map(c => typeof c === 'string' ? c : (c.name || c.originalName || c.header));
  const columnSpec = (spec, name) => (spec.columns || []).find(c => String(typeof c === 'string' ? c : (c.name || c.originalName || c.header)).toLowerCase() === String(name).toLowerCase());
  function walk(e) { if (!e) return null; if (e.type === 'column') { let owner; if (e.table) owner = aliases[String(e.table).toLowerCase()]; else { const matches = Object.entries(aliases).filter(([, candidate]) => columnSpec(candidate.spec, e.name)); if (!matches.length) return reject(`Unknown column ${e.name}`, 0, `Unknown column ${e.name}`); if (matches.length > 1) { const names = matches.map(([alias]) => alias); return reject(`Ambiguous column ${e.name}`, 0, `Column ${e.name} exists on multiple tables (${names.join(', ')}) in this query; qualify it with a table alias, e.g. ${names[0]}.${e.name}`); } e.table = matches[0][0]; owner = matches[0][1]; } if (!owner) return reject(`Unknown table alias ${e.table}`, 0, `Unknown table alias ${e.table}`); const column = columnSpec(owner.spec, e.name); if (!column) return reject(`Unknown column ${e.name}`, 0, `Unknown column ${e.name}`); e.resolvedTable = owner.key; e.resolvedAlias = e.table || ast.from.alias; e.resolvedHeader = typeof column === 'string' ? e.name : (column.header || column.displayHeader || e.name); e.sourceHeader = e.resolvedHeader; const joined = ast.joins.find(j => String(j.table.alias).toLowerCase() === String(e.resolvedAlias).toLowerCase()); if (joined && columnSpec(aliases[String(ast.from.alias).toLowerCase()].spec, e.name)) e.resolvedHeader = `${joined.table.alias}.${e.resolvedHeader}`; } if (e.type === 'call') { if (!['SUM','COUNT','AVG','MIN','MAX','COALESCE','STRFTIME','DATE','JULIANDAY','DISTINCT','ROUND','ABS','PRINTF'].includes(e.name)) return reject(`${e.name} function`, 0, `${e.name} is not supported`); const r = validateCall(e); if (r) return r; } for (const v of Object.values(e)) { if (Array.isArray(v)) for (const x of v) { if (x && x.type) { const r = walk(x); if (r) return r; } else if (x && typeof x === 'object') { const r = walk(x.when) || walk(x.then); if (r) return r; } } else if (v && v.type) { const r = walk(v); if (r) return r; } } return null; }
  function validateCall(e) { const literal = (x, value) => x && x.type === 'literal' && x.literalType === 'string' && (value == null || String(x.value).toLowerCase() === value); const numeric = x => x && !(x.type === 'literal' && x.literalType === 'string'); if (e.name === 'STRFTIME') { if (e.args.length !== 2 || !literal(e.args[0])) return reject('STRFTIME arguments', 0, 'STRFTIME requires a string format and one date expression'); const bad = String(e.args[0].value).match(/%[^Ymdw]/); if (bad) return reject(`STRFTIME token ${bad[0]}`, 0, `${bad[0]} is not supported in STRFTIME; only %Y, %m, %d, and %w are supported`); } if (e.name === 'DATE') { if ((e.args.length !== 1 && e.args.length !== 2) || !literal(e.args[0], 'now')) return reject('DATE arguments', 0, "DATE supports date('now') and date('now', '+N days|months|years')"); const modifier = e.args[1] && String(e.args[1].value).toLowerCase(); if (e.args.length === 2 && (!literal(e.args[1]) || (!/^[+-]\d+\s+(day|days|month|months|year|years)$/i.test(e.args[1].value) && !['start of month', 'start of year', 'start of day', 'start of week'].includes(modifier)))) return reject('DATE offset', 0, "DATE offset must be a signed day, month, or year duration, or a supported start-of-period modifier"); } if (e.name === 'JULIANDAY' && e.args.length !== 1) return reject('JULIANDAY arguments', 0, 'JULIANDAY requires one date expression'); if (e.name === 'COALESCE' && !e.args.length) return reject('COALESCE arguments', 0, 'COALESCE requires at least one expression'); if (e.name === 'ROUND' && ((e.args.length !== 1 && e.args.length !== 2) || !e.args.every(numeric))) return reject('ROUND arguments', 0, 'ROUND requires one or two numeric expressions'); if (e.name === 'ABS' && (e.args.length !== 1 || !numeric(e.args[0]))) return reject('ABS arguments', 0, 'ABS requires one numeric expression'); if (e.name === 'PRINTF' && (e.args.length !== 2 || !literal(e.args[0]) || !/^%0(\d+)d$/.test(e.args[0].value))) return reject('PRINTF arguments', 0, 'PRINTF requires a literal %0Nd format and one expression'); if (['SUM','AVG','MIN','MAX','DISTINCT'].includes(e.name) && e.args.length !== 1) return reject(`${e.name} arguments`, 0, `${e.name} requires one expression`); if (['SUM','AVG','MIN','MAX'].includes(e.name) && (e.args[0].type === 'star' || containsAggregate(e.args[0]))) return reject(`${e.name} argument shape`, 0, `${e.name}(*) is not supported; and ${e.name} cannot itself contain another aggregate`, `${e.name} accepts a column, expression (arithmetic, function calls, CAST, etc.), or CASE — but not * and not a nested aggregate. If you're seeing this, check for one of those two specific cases; you can also pre-compute the value in a subquery column and aggregate that column in the outer query if you prefer staging.`); if (e.name === 'COUNT' && (e.args.length !== 1 || (e.args[0].type !== 'star' && e.args[0].type !== 'column' && !(e.args[0].type === 'call' && e.args[0].name === 'DISTINCT')))) return reject('COUNT arguments', 0, 'COUNT requires *, a column, or DISTINCT column'); if (e.name === 'COUNT' && e.args[0].type === 'call' && e.args[0].name === 'DISTINCT' && (!e.args[0].args[0] || e.args[0].args[0].type === 'star' || containsAggregate(e.args[0].args[0]))) return reject('COUNT(DISTINCT ...) argument shape', 0, 'COUNT(DISTINCT ...) requires an expression; * and nested aggregates are not supported', 'COUNT(DISTINCT ...) accepts a column, expression (including CASE), or other non-aggregate expression — but not a missing argument, * or a nested aggregate. If you prefer staging, pre-compute the value in a subquery column and COUNT(DISTINCT that_column) in the outer query.'); return null; }
  for (let i = 0; i < ast.select.length; i++) { const item = ast.select[i]; if (item.type === 'star') { const wanted = ((report && report.columns) || []).map(c => typeof c === 'string' ? c : c.name); const actual = cols(tables[baseKey]); if (!wanted.length || wanted.length !== actual.length || wanted.some(x => !actual.includes(x))) return reject('SELECT *', item.position, 'SELECT * is not supported unless it exactly matches report.columns', 'List columns explicitly with aliases matching report.columns[].name.');
    // A star item has no per-column AST node for the planner/emitter to walk; expand it in
    // place into the same {expression, alias, position} select items an equivalent
    // hand-written SELECT would produce from parser.js's item().
    const expanded = wanted.map(name => ({ expression: { type: 'column', table: ast.from.alias, name }, alias: name, position: item.position }));
    for (const col of expanded) { const r = walk(col.expression); if (r) return r; }
    ast.select.splice(i, 1, ...expanded); i += expanded.length - 1;
  } else { const r = walk(item.expression); if (r) return r; } }
  function resolveSelectAlias(e, groupBy) {
    if (!e || e.type !== 'column' || e.table) return null;
    const item = ast.select.find(x => x.type !== 'star' && String(x.alias).toLowerCase() === String(e.name).toLowerCase());
    if (!item) return null;
    if (groupBy && containsAggregate(item.expression)) return reject('GROUP BY aggregate alias', 0, `GROUP BY cannot reference an aggregate alias like "${e.name}"`, 'Add the column to GROUP BY or wrap it in an aggregate; GROUP BY names columns, not aggregate aliases.');
    return item.expression;
  }
  for (let i = 0; i < (ast.groupBy || []).length; i++) { const resolved = resolveSelectAlias(ast.groupBy[i], true); if (resolved && resolved.ok === false) return resolved; if (resolved) ast.groupBy[i] = resolved; else { const r = walk(ast.groupBy[i]); if (r) return r; } }
  // HAVING, unlike WHERE, can refer to SELECT-list aliases. Resolve aliases
  // throughout its boolean expression before validating the underlying columns.
  function resolveHavingAliases(e) {
    if (!e || typeof e !== 'object') return e;
    const resolved = resolveSelectAlias(e, false);
    if (resolved) return resolved;
    for (const [name, value] of Object.entries(e)) {
      if (Array.isArray(value)) e[name] = value.map(resolveHavingAliases);
      else if (value && typeof value === 'object') e[name] = resolveHavingAliases(value);
    }
    return e;
  }
  if (ast.having) ast.having = resolveHavingAliases(ast.having);
  for (const e of [ast.where, ast.having, ...ast.joins.map(x => x.on)]) { const r = walk(e); if (r) return r; }
  for (let i = 0; i < (ast.orderBy || []).length; i++) { const resolved = resolveSelectAlias(ast.orderBy[i].expression, false); if (resolved) ast.orderBy[i].expression = resolved; else { const r = walk(ast.orderBy[i].expression); if (r) return r; } }
  const hasAggregate = ast.select.some(x => x.expression && containsAggregate(x.expression));
  if (hasAggregate && ast.joins.length >= 2) { const risky = ast.select.some(item => { const refs = []; const scan = e => { if (!e || typeof e !== 'object') return; if (e.type === 'column') refs.push(String(e.resolvedAlias || e.table).toLowerCase()); else Object.values(e).forEach(v => Array.isArray(v) ? v.forEach(scan) : scan(v)); }; if (containsAggregate(item.expression)) scan(item.expression); return refs.some(a => a !== String(ast.from.alias).toLowerCase()); }); if (risky) return reject('join fan-out', ast.joins[1].table.position, 'Aggregation over columns from 2+ joined child tables can double-count rows through join fan-out; this mixed child aggregate shape is not safe to rewrite automatically.', 'Allowed shapes: (1) put the measure table in FROM and only join N:1 lookup tables, aggregating base columns; (2) one raw child join + aggregate is fine — two are not; (3) pre-aggregate each child in a subquery (SELECT parent_id, SUM(x) ... GROUP BY parent_id), join those, then wrap everything in an outer SELECT ... FROM (...) t GROUP BY so aggregates read only t.*. See docs/excel-report-sql-patterns.md. If none of these shapes fit, this is a compiler-capability gap — try docs/excel-report-sql-patterns.md first; queries.excel_power_query (hand-written M) should only be used as a genuine last resort, since small/weaker authoring models reliably cannot hand-write correct Power Query M.'); }
  // A known, single-column primary key functionally determines the other columns
  // of its own table. Add those projected columns as explicit group keys so the
  // downstream plan has an unambiguous representation of that SQL rule.
  const columnIdentity = e => e && e.type === 'column' ? `${String(e.resolvedAlias).toLowerCase()}\u0000${String(e.name).toLowerCase()}` : null;
  const groupedColumns = new Set((ast.groupBy || []).map(columnIdentity).filter(Boolean));
  for (const group of [...(ast.groupBy || [])]) {
    if (!group || group.type !== 'column') continue;
    const spec = aliases[String(group.resolvedAlias).toLowerCase()] && aliases[String(group.resolvedAlias).toLowerCase()].spec;
    const primaryKeys = (spec && spec.columns || []).filter(c => c && typeof c === 'object' && c.isPrimaryKey === true);
    const groupedColumn = columnSpec(spec || {}, group.name);
    if (primaryKeys.length !== 1 || groupedColumn !== primaryKeys[0]) continue;
    for (const item of ast.select) {
      const e = item.expression;
      if (!e || containsAggregate(e) || e.type !== 'column' || e.resolvedTable !== group.resolvedTable || e.resolvedAlias !== group.resolvedAlias || groupedColumns.has(columnIdentity(e))) continue;
      ast.groupBy.push({ type: 'column', table: e.resolvedAlias, name: e.name, resolvedTable: e.resolvedTable, resolvedAlias: e.resolvedAlias, resolvedHeader: e.resolvedHeader, sourceHeader: e.sourceHeader });
      groupedColumns.add(columnIdentity(e));
    }
  }
  const grouped = new Set((ast.groupBy || []).map(key));
  const isKnownSinglePrimaryKey = e => {
    if (!e || e.type !== 'column') return false;
    const spec = aliases[String(e.resolvedAlias).toLowerCase()] && aliases[String(e.resolvedAlias).toLowerCase()].spec;
    const primaryKeys = (spec && spec.columns || []).filter(c => c && typeof c === 'object' && c.isPrimaryKey === true);
    return primaryKeys.length === 1 && columnSpec(spec || {}, e.name) === primaryKeys[0];
  };
  const onlyGroupedColumns = e => {
    const refs = [];
    const scan = value => { if (!value || typeof value !== 'object') return; if (value.type === 'column') { refs.push(value); return; } Object.values(value).forEach(v => Array.isArray(v) ? v.forEach(scan) : scan(v)); };
    scan(e);
    return refs.length > 0 && refs.every(ref => groupedColumns.has(columnIdentity(ref)));
  };
  if (hasAggregate) for (const item of ast.select) if (item.expression && !containsAggregate(item.expression) && !grouped.has(key(item.expression)) && !onlyGroupedColumns(item.expression)) return reject('GROUP BY projection consistency', item.position, 'Every non-aggregated projected column must appear in GROUP BY', 'Add the column to GROUP BY or wrap it in an aggregate; GROUP BY names columns, not aggregate aliases.');
  for (const g of ast.groupBy || []) if (!isKnownSinglePrimaryKey(g) && !ast.select.some(x => x.expression && (key(x.expression) === key(g) || (onlyGroupedColumns(x.expression) && columnIdentity(g) && expressionContainsColumn(x.expression, columnIdentity(g)))))) return reject('GROUP BY non-projected expression', 0, 'GROUP BY on a non-projected expression is not supported');
  return { ok: true, ast, aliases };
}
function containsAggregate(e) { if (!e || typeof e !== 'object') return false; if (e.type === 'call' && aggregateNames.has(e.name)) return true; return Object.values(e).some(v => Array.isArray(v) ? v.some(containsAggregate) : containsAggregate(v)); }
function expressionContainsColumn(e, identity) { if (!e || typeof e !== 'object') return false; if (e.type === 'column') return `${String(e.resolvedAlias).toLowerCase()}\u0000${String(e.name).toLowerCase()}` === identity; return Object.values(e).some(v => Array.isArray(v) ? v.some(x => expressionContainsColumn(x, identity)) : expressionContainsColumn(v, identity)); }
function key(e) { return JSON.stringify(e, (k, v) => ['resolvedTable','resolvedAlias','resolvedHeader'].includes(k) ? undefined : v); }
module.exports = { validate, containsAggregate };
