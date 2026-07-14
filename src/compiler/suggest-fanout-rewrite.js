'use strict';
// Authoring-time helper for the "join fan-out" rejection in validate.js.
//
// That rejection is correct: aggregating columns from 2+ raw joined 1:N
// branches over the SAME base row can double-count through a naive sequential
// join. The fix (see docs/excel-report-sql-patterns.md, Pattern D) is for the
// SQL author to pre-aggregate each branch into its own subquery, join those,
// then wrap the result in an outer SELECT ... GROUP BY. This module attempts
// that reshape automatically and hands back plain SQL text for a human to
// review and paste into report.query - it never compiles anything itself and
// never touches the real compiler (parser.js/validate.js/plan.js/emit.js).
// If the shape is anything more complex than the well-understood cases below,
// it bails with a reason instead of guessing.
const { parseSql } = require('./parser');
const { containsAggregate } = require('./validate');

const AGG_FUNCS = new Set(['SUM', 'COUNT', 'AVG', 'MIN', 'MAX']);
const bail = reason => ({ ok: false, reason });

function forEachExprChild(e, visit) {
  switch (e.type) {
    case 'unary': case 'isnull': visit(e.expression); break;
    case 'binary': visit(e.left); visit(e.right); break;
    case 'in': visit(e.expression); e.values.forEach(visit); break;
    case 'between': visit(e.expression); visit(e.low); visit(e.high); break;
    case 'cast': visit(e.value); break;
    case 'call': e.args.forEach(visit); break;
    case 'case': e.branches.forEach(b => { visit(b.when); visit(b.then); }); visit(e.otherwise); break;
    default: break;
  }
}
function forEachColumn(e, visit) {
  if (!e || typeof e !== 'object') return;
  if (e.type === 'column') { visit(e); return; }
  forEachExprChild(e, x => forEachColumn(x, visit));
}
function findAggregateCalls(e, out = []) {
  if (!e || typeof e !== 'object') return out;
  if (e.type === 'call' && AGG_FUNCS.has(e.name)) { out.push(e); return out; }
  forEachExprChild(e, x => findAggregateCalls(x, out));
  return out;
}
// Pre-order: fn may return a replacement node to stop recursion at this node.
function mapExpr(e, fn) {
  if (!e || typeof e !== 'object') return e;
  const replacement = fn(e);
  if (replacement) return replacement;
  switch (e.type) {
    case 'unary': case 'isnull': return { ...e, expression: mapExpr(e.expression, fn) };
    case 'binary': return { ...e, left: mapExpr(e.left, fn), right: mapExpr(e.right, fn) };
    case 'in': return { ...e, expression: mapExpr(e.expression, fn), values: e.values.map(v => mapExpr(v, fn)) };
    case 'between': return { ...e, expression: mapExpr(e.expression, fn), low: mapExpr(e.low, fn), high: mapExpr(e.high, fn) };
    case 'cast': return { ...e, value: mapExpr(e.value, fn) };
    case 'call': return { ...e, args: e.args.map(a => mapExpr(a, fn)) };
    case 'case': return { ...e, branches: e.branches.map(b => ({ when: mapExpr(b.when, fn), then: mapExpr(b.then, fn) })), otherwise: mapExpr(e.otherwise, fn) };
    default: return e;
  }
}
const structuralKey = e => JSON.stringify(e);

function printExpr(e) {
  switch (e.type) {
    case 'star': return '*';
    case 'column': return e.table ? `${e.table}.${e.name}` : e.name;
    case 'literal':
      if (e.literalType === 'string') return `'${String(e.value).replace(/'/g, "''")}'`;
      if (e.literalType === 'null') return 'NULL';
      return String(e.value);
    case 'unary': return e.op === 'NOT' ? `NOT ${printExpr(e.expression)}` : `${e.op}${printExpr(e.expression)}`;
    case 'binary': return `${printExpr(e.left)} ${e.op} ${printExpr(e.right)}`;
    case 'isnull': return `${printExpr(e.expression)} IS ${e.not ? 'NOT ' : ''}NULL`;
    case 'in': return `${printExpr(e.expression)} ${e.negated ? 'NOT ' : ''}IN (${e.values.map(printExpr).join(', ')})`;
    case 'between': return `${printExpr(e.expression)} BETWEEN ${printExpr(e.low)} AND ${printExpr(e.high)}`;
    case 'cast': return `CAST(${printExpr(e.value)} AS ${e.asType})`;
    case 'call': {
      if (e.name === 'DISTINCT') return `DISTINCT ${printExpr(e.args[0])}`;
      const args = e.args.map(a => a.type === 'star' ? '*' : (a.type === 'call' && a.name === 'DISTINCT' ? printExpr(a) : printExpr(a)));
      return `${e.name}(${args.join(', ')})`;
    }
    case 'case': return `CASE ${e.branches.map(b => `WHEN ${printExpr(b.when)} THEN ${printExpr(b.then)}`).join(' ')} ELSE ${printExpr(e.otherwise)} END`;
    default: throw new Error(`printExpr: unsupported node type ${e.type}`);
  }
}

function equalityRefs(on) {
  const equalities = []; const remainders = [];
  const visit = node => {
    if (!node) return;
    if (node.type === 'binary' && node.op === 'AND') { visit(node.left); visit(node.right); return; }
    if (node.type === 'binary' && node.op === '=' && node.left && node.left.type === 'column' && node.right && node.right.type === 'column') { equalities.push(node); return; }
    remainders.push(node);
  };
  visit(on);
  return { equalities, remainders };
}

function suggestFanoutRewrite(sql) {
  const parsed = parseSql(sql, {});
  if (!parsed.ok) return bail(`SQL does not parse: ${(parsed.rejection && parsed.rejection.message) || 'unknown error'}`);
  const ast = parsed.ast;
  if (ast.type === 'union') return bail('UNION queries are not supported by the auto-suggestion tool.');
  if (ast.from.subquery) return bail('FROM is already a subquery; reshape this by hand (see docs/excel-report-sql-patterns.md).');
  if (ast.joins.some(j => j.table.subquery)) return bail('A JOIN target is already a subquery; reshape this by hand.');
  if (ast.joins.length < 2) return bail('Fewer than 2 joins; this is not the join fan-out shape validate.js rejects.');
  if (ast.select.some(x => x.type === 'star')) return bail('SELECT * is not supported by the auto-suggestion tool.');
  if (!ast.select.some(x => containsAggregate(x.expression))) return bail('No aggregate in SELECT; this is not the join fan-out shape.');

  const baseAlias = ast.from.alias;
  const baseLower = String(baseAlias).toLowerCase();

  let hasUnqualified = false;
  const markQualified = e => forEachColumn(e, c => { if (!c.table) hasUnqualified = true; });
  ast.select.forEach(x => markQualified(x.expression));
  markQualified(ast.where); (ast.groupBy || []).forEach(markQualified); markQualified(ast.having);
  (ast.orderBy || []).forEach(x => markQualified(x.expression));
  ast.joins.forEach(j => markQualified(j.on));
  if (hasUnqualified) return bail('Query has unqualified column references; qualify every column with its table alias before auto-suggestion can analyze it.');

  // Resolve each join's correlation to a single "parent" alias (base or another join).
  const parentInfo = new Map();
  const childrenOf = new Map();
  for (const j of ast.joins) {
    const ownAlias = String(j.table.alias).toLowerCase();
    if (j.kind === 'CROSS') { parentInfo.set(ownAlias, 'AMBIGUOUS'); continue; }
    const { equalities, remainders } = equalityRefs(j.on);
    const outer = equalities.filter(eq => (String(eq.left.table).toLowerCase() === ownAlias) !== (String(eq.right.table).toLowerCase() === ownAlias));
    const otherAliases = new Set(outer.map(eq => String(eq.left.table).toLowerCase() === ownAlias ? String(eq.right.table).toLowerCase() : String(eq.left.table).toLowerCase()));
    if (outer.length !== 1 || otherAliases.size !== 1) { parentInfo.set(ownAlias, 'AMBIGUOUS'); continue; }
    const parentAlias = [...otherAliases][0];
    const eq = outer[0];
    const ownColumn = String(eq.left.table).toLowerCase() === ownAlias ? eq.left : eq.right;
    const parentColumn = String(eq.left.table).toLowerCase() === ownAlias ? eq.right : eq.left;
    parentInfo.set(ownAlias, { parentAlias, ownColumn, parentColumn, remainders, join: j });
    if (!childrenOf.has(parentAlias)) childrenOf.set(parentAlias, []);
    childrenOf.get(parentAlias).push(j);
  }

  // Which aliases does a select-list aggregate call actually read from?
  const riskyAliasSet = new Set();
  for (const item of ast.select) {
    for (const call of findAggregateCalls(item.expression)) {
      const aliases = new Set();
      call.args.forEach(a => forEachColumn(a, c => aliases.add(String(c.table).toLowerCase())));
      aliases.delete(baseLower);
      if (aliases.size > 1) return bail('An aggregate expression mixes columns from two different joined branches; not supported by auto-suggestion.');
      if (aliases.size === 1) riskyAliasSet.add([...aliases][0]);
    }
  }
  if (!riskyAliasSet.size) return bail('No aggregate references a joined branch; this SQL was not rejected for join fan-out.');
  for (const alias of riskyAliasSet) {
    const info = parentInfo.get(alias);
    if (info === 'AMBIGUOUS' || !info) return bail(`The join for "${alias}" has a compound or cross-alias correlation key; not supported by auto-suggestion.`);
  }

  // Fold descendant joins that exist purely to filter/lookup for a risky branch
  // (never referenced anywhere except inside that branch's own JOIN..ON).
  const referencedElsewhere = alias => {
    let found = false;
    const scan = e => forEachColumn(e, c => { if (String(c.table).toLowerCase() === alias) found = true; });
    ast.select.forEach(x => scan(x.expression));
    scan(ast.where); (ast.groupBy || []).forEach(scan); scan(ast.having);
    (ast.orderBy || []).forEach(x => scan(x.expression));
    for (const j of ast.joins) { const own = String(j.table.alias).toLowerCase(); if (own !== alias) scan(j.on); }
    return found;
  };
  const foldedSet = new Set();
  for (const risky of riskyAliasSet) {
    const queue = [...(childrenOf.get(risky) || [])];
    while (queue.length) {
      const child = queue.shift();
      const childAlias = String(child.table.alias).toLowerCase();
      if (foldedSet.has(childAlias)) continue;
      if (riskyAliasSet.has(childAlias)) return bail(`A branch join ("${childAlias}") is itself aggregated inside another branch; not supported by auto-suggestion.`);
      if (referencedElsewhere(childAlias)) return bail(`Join "${childAlias}" is used outside its own JOIN condition; not supported by auto-suggestion.`);
      foldedSet.add(childAlias);
      queue.push(...(childrenOf.get(childAlias) || []));
    }
  }

  // Assign a synthetic-or-natural output name to each distinct aggregate call over a risky branch.
  const callAssignments = new Map(); // structuralKey -> { name, outerFn, branchAlias, remaindersChecked }
  let synth = 1;
  const pickName = k => {
    for (const item of ast.select) {
      const calls = findAggregateCalls(item.expression);
      if (calls.length === 1 && structuralKey(calls[0]) === k) return item.alias;
    }
    return null;
  };
  for (const item of ast.select) {
    for (const call of findAggregateCalls(item.expression)) {
      const aliases = new Set();
      call.args.forEach(a => forEachColumn(a, c => aliases.add(String(c.table).toLowerCase())));
      aliases.delete(baseLower);
      if (aliases.size !== 1) continue;
      const branchAlias = [...aliases][0];
      if (!riskyAliasSet.has(branchAlias)) continue;
      if (call.name === 'AVG') return bail(`AVG over joined branch "${branchAlias}" cannot be safely decomposed by auto-suggestion; compute SUM and COUNT separately and divide by hand.`);
      const k = structuralKey(call);
      if (!callAssignments.has(k)) {
        const outerFn = call.name === 'COUNT' ? 'SUM' : call.name;
        callAssignments.set(k, { name: pickName(k) || `_agg${synth++}`, outerFn, branchAlias, call });
      }
    }
  }

  // Build each branch's own pre-aggregate subquery.
  const branches = new Map(); // alias -> { table, fkColumn, remainder, aggregates: [{call,name,outerFn}], foldedJoins: [join] }
  for (const alias of riskyAliasSet) {
    const info = parentInfo.get(alias);
    let remainderExpr = info.remainders.length ? info.remainders.reduce((a, b) => a ? { type: 'binary', op: 'AND', left: a, right: b } : b, null) : null;
    if (remainderExpr) {
      let bad = false;
      forEachColumn(remainderExpr, c => { if (String(c.table).toLowerCase() !== alias) bad = true; });
      if (bad) return bail(`Join condition for "${alias}" filters using a column from another table; not supported by auto-suggestion.`);
    }
    const foldedJoins = ast.joins.filter(j => foldedSet.has(String(j.table.alias).toLowerCase()) && (() => {
      // only include descendants that trace back to this branch
      let cur = String(j.table.alias).toLowerCase();
      while (true) {
        const p = parentInfo.get(cur);
        if (!p || p === 'AMBIGUOUS') return false;
        if (p.parentAlias === alias) return true;
        if (!riskyAliasSet.has(p.parentAlias) && !foldedSet.has(p.parentAlias)) return false;
        cur = p.parentAlias;
      }
    })());
    branches.set(alias, {
      table: info.join.table.name, alias, fkColumn: info.ownColumn.name, parentAlias: info.parentAlias,
      parentColumn: info.parentColumn.name, kind: info.join.kind, remainderExpr, foldedJoins, aggregates: [],
    });
  }
  for (const [, assign] of callAssignments) {
    const branch = branches.get(assign.branchAlias);
    branch.aggregates.push(assign);
  }

  // Column-naming for every "leftover" plain column that must pass through staging untouched.
  const columnNaming = new Map(); // "alias\0name" -> staging output name
  const usedNames = new Set(Array.from(callAssignments.values()).map(a => a.name));
  const registerColumn = c => {
    const alias = String(c.table).toLowerCase();
    if (riskyAliasSet.has(alias) || foldedSet.has(alias)) return; // consumed inside a branch aggregate only
    const key = `${alias} ${c.name.toLowerCase()}`;
    if (columnNaming.has(key)) return;
    let name = c.name;
    if (usedNames.has(name)) name = `${alias}_${c.name}`;
    usedNames.add(name);
    columnNaming.set(key, name);
  };
  const scanForColumns = e => forEachColumn(e, registerColumn);
  ast.select.forEach(x => scanForColumns(x.expression));
  scanForColumns(ast.where); (ast.groupBy || []).forEach(scanForColumns); scanForColumns(ast.having);
  (ast.orderBy || []).forEach(x => scanForColumns(x.expression));

  // Safety net: after consuming aggregate calls, no risky/folded-alias column should remain anywhere.
  const outerSubstitute = e => e && mapExpr(e, node => {
    if (node.type === 'call' && AGG_FUNCS.has(node.name)) {
      const assign = callAssignments.get(structuralKey(node));
      if (assign) return { type: 'call', name: assign.outerFn, args: [{ type: 'column', table: 't', name: assign.name }] };
      return null;
    }
    if (node.type === 'column') {
      const alias = String(node.table).toLowerCase();
      if (riskyAliasSet.has(alias) || foldedSet.has(alias)) return undefined; // leave unresolved -> caught below
      const staged = columnNaming.get(`${alias} ${node.name.toLowerCase()}`);
      return { type: 'column', table: 't', name: staged || node.name };
    }
    return null;
  });
  const outerSelectItems = ast.select.map(item => ({ alias: item.alias, expression: outerSubstitute(item.expression) }));
  const outerHaving = ast.having ? outerSubstitute(ast.having) : null;
  const outerOrderBy = (ast.orderBy || []).map(o => ({ expression: outerSubstitute(o.expression), direction: o.direction }));
  let leftoverRisky = false;
  const checkLeftover = e => forEachColumn(e, c => { const a = String(c.table).toLowerCase(); if (riskyAliasSet.has(a) || foldedSet.has(a)) leftoverRisky = true; });
  outerSelectItems.forEach(x => checkLeftover(x.expression)); if (outerHaving) checkLeftover(outerHaving); outerOrderBy.forEach(x => checkLeftover(x.expression));
  if (leftoverRisky) return bail('A joined branch column is used in a way auto-suggestion cannot safely relocate (outside a plain aggregate); not supported.');

  if (ast.where) { let bad = false; forEachColumn(ast.where, c => { const a = String(c.table).toLowerCase(); if (riskyAliasSet.has(a) || foldedSet.has(a)) bad = true; }); if (bad) return bail('WHERE filters on a joined branch column; moving that filter safely requires a manual Pattern-D rewrite.'); }

  // --- Emit SQL text ---
  const printJoinOn = j => `${j.kind} JOIN ${j.table.name} ${j.table.alias} ON ${printExpr(j.on)}`;

  const branchSql = branch => {
    const lines = [`SELECT ${branch.alias}.${branch.fkColumn}`];
    for (const agg of branch.aggregates) lines[0] += `, ${printExpr(agg.call)} AS ${agg.name}`;
    lines.push(`FROM ${branch.table} ${branch.alias}`);
    for (const j of branch.foldedJoins) lines.push(printJoinOn(j));
    if (branch.remainderExpr) lines.push(`WHERE ${printExpr(branch.remainderExpr)}`);
    lines.push(`GROUP BY ${branch.alias}.${branch.fkColumn}`);
    return `(${lines.join('\n  ')})`;
  };

  const stagingJoinLines = [];
  for (const j of ast.joins) {
    const alias = String(j.table.alias).toLowerCase();
    if (foldedSet.has(alias)) continue;
    if (riskyAliasSet.has(alias)) {
      const branch = branches.get(alias);
      stagingJoinLines.push(`LEFT JOIN ${branchSql(branch)} ${branch.alias} ON ${branch.parentAlias}.${branch.parentColumn} = ${branch.alias}.${branch.fkColumn}`);
    } else {
      stagingJoinLines.push(printJoinOn(j));
    }
  }
  const stagingSelectItems = [];
  for (const [key, name] of columnNaming) { const [alias, col] = key.split(' '); stagingSelectItems.push(`${alias}.${col} AS ${name}`); }
  for (const [, assign] of callAssignments) stagingSelectItems.push(`${assign.branchAlias}.${assign.name} AS ${assign.name}`);
  if (!stagingSelectItems.length) return bail('Nothing to project into the staging query; not supported by auto-suggestion.');

  const stagingLines = [`SELECT ${stagingSelectItems.join(', ')}`, `FROM ${ast.from.name} ${baseAlias}`, ...stagingJoinLines];
  if (ast.where) stagingLines.push(`WHERE ${printExpr(ast.where)}`);

  // The staging FROM is a subquery, so it carries no primary-key metadata: the
  // real compiler's "GROUP BY on a non-projected column is OK if it's a known
  // single primary key" exception (validate.js) never fires for it. Every
  // GROUP BY entry must therefore be one of the outer SELECT's own projected,
  // non-aggregate columns - the original GROUP BY key (often a bare PK that
  // relied on that exception) is redundant with those once translated, since
  // they all come from the same source row, so it is dropped rather than
  // carried over literally.
  const outerSelectPrinted = outerSelectItems.map(x => `${printExpr(x.expression)} AS ${x.alias}`);
  const requiredGroupBy = new Set();
  for (const item of outerSelectItems) if (!findAggregateCalls(item.expression).length) forEachColumn(item.expression, c => requiredGroupBy.add(`${c.table}.${c.name}`));
  const outerGroupBy = [...requiredGroupBy];

  const outerLines = [`SELECT ${outerSelectPrinted.join(', ')}`, `FROM (`, ...stagingLines.map(l => `  ${l}`), `) t`];
  if (outerGroupBy.length) outerLines.push(`GROUP BY ${outerGroupBy.join(', ')}`);
  if (outerHaving) outerLines.push(`HAVING ${printExpr(outerHaving)}`);
  if (outerOrderBy.length) outerLines.push(`ORDER BY ${outerOrderBy.map(o => `${printExpr(o.expression)} ${o.direction}`).join(', ')}`);

  const notes = [];
  if ([...callAssignments.values()].some(a => a.outerFn === 'SUM' && a.call.name === 'COUNT')) notes.push('A COUNT over a joined branch was rewritten as an outer SUM of per-key counts (the mathematically correct composition).');
  if (usedNames.size !== new Set(usedNames).size) notes.push('Some staging column names were qualified with their table alias to avoid collisions.');
  notes.push('Generated by suggest-fanout-rewrite.js: review before pasting into report.query, then re-run the real compiler to confirm it compiles.');

  return { ok: true, sql: outerLines.join('\n'), notes };
}

module.exports = { suggestFanoutRewrite };
