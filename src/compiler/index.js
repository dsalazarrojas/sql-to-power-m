'use strict';
const { parseSql } = require('./parser'); const { rewriteExists } = require('./rewrite-exists'); const { validate } = require('./validate'); const { makePlan } = require('./plan'); const { emitM } = require('./emit'); const { powerQueryStructuralDiagnostics, findPowerQueryFunctionCalls, splitTopLevelArgs } = require('./power-query-lint');
const { suggestFanoutRewrite } = require('./suggest-fanout-rewrite');
function duplicateRenameTargetDiagnostics(mCode) {
  const diagnostics = [];
  for (const call of findPowerQueryFunctionCalls(mCode, 'Table.RenameColumns')) {
    const pairs = splitTopLevelArgs(call.args)[1] || '';
    const targets = new Set(); let match;
    const pair = /\{\s*"((?:""|[^"])*)"\s*,\s*"((?:""|[^"])*)"\s*\}/g;
    while ((match = pair.exec(pairs))) {
      const target = match[2].replace(/""/g, '"');
      if (targets.has(target)) diagnostics.push(`Table.RenameColumns has duplicate target column name "${target}" in one call.`);
      targets.add(target);
    }
  }
  return diagnostics;
}
function compileSqlToM({ sql, report, headerContext }) { const parsed = parseSql(sql, { report }); if (!parsed.ok) return { ok: false, mCode: null, plan: null, rejections: [parsed.rejection] }; const rewritten = rewriteExists(parsed.ast); if (!rewritten.ok) return { ok: false, mCode: null, plan: null, rejections: [rewritten.rejection] }; const checked = validate(parsed.ast, headerContext, report); if (!checked.ok) return { ok: false, mCode: null, plan: null, rejections: [checked.rejection] }; let plan; try { plan = makePlan(checked.ast, headerContext, report); } catch (err) { if (err && err.rejection) return { ok: false, mCode: null, plan: null, rejections: [err.rejection] }; throw err; } const mCode = emitM(plan, report); const nested = parsed.ast.type === 'union' || parsed.ast.from && parsed.ast.from.subquery || parsed.ast.joins && parsed.ast.joins.some(j => j.table.subquery); const diagnostics = (nested ? [] : powerQueryStructuralDiagnostics(mCode)).concat(duplicateRenameTargetDiagnostics(mCode)); if (diagnostics.length) return { ok: false, mCode: null, plan: null, rejections: [{ code: 'internal-error', construct: 'M emission', position: 0, message: diagnostics.join('; '), hint: 'This is a compiler bug; please report the SQL and diagnostic.' }] }; return { ok: true, mCode, plan, rejections: [] }; }
module.exports = { compileSqlToM, suggestFanoutRewrite };
