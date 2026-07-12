'use strict';

// Nested subquery joins compile to a `(let ... in ...)` expression embedded as a call
// argument (e.g. the right side of Table.NestedJoin). That inner let is its own lexical
// scope with its own step names, which commonly reuse the outer scope's names (Source,
// FilteredRows, ...). All structural checks below must resolve a call's step references
// against the scope that actually encloses that call, not a single flat, whole-document
// step/header map, or same-named steps in sibling/nested scopes collide.
function powerQueryStructuralDiagnostics(mCode) {
  const source = String(mCode || '');
  const diagnostics = [];
  if (!/^\s*let\b/i.test(source) || !/\bin\b/i.test(source)) {
    diagnostics.push('Power Query M must use a complete let ... in ... expression.');
    return diagnostics;
  }
  const rootScope = buildLetScopeTree(source, 0, source.length);
  if (!rootScope) {
    diagnostics.push('Power Query M must use a complete let ... in ... expression.');
    return diagnostics;
  }
  walkLetScopeTree(rootScope, (scope, isRoot) => {
    for (const step of scope.assignments) {
      const chained = String(step.expression || '').match(/^\s*(#?"(?:""|[^"])+"|[A-Za-z_][A-Za-z0-9_]*)\s*=/);
      if (chained) {
        diagnostics.push(`Power Query step "${step.name}" contains a chained assignment ("${step.name} = ${normalizePowerQueryIdentifier(chained[1])} = ..."). M let steps must be one assignment per step.`);
      }
    }
    if (isRoot) {
      if (!scope.inTarget) {
        diagnostics.push('Power Query M must end with "in <stepName>" so Excel can resolve the final step.');
      } else if (!scope.ownSteps.has(scope.inTarget.toLowerCase())) {
        diagnostics.push(`Power Query final step "${scope.inTarget}" is not defined in the let block.`);
      }
    }
  });
  for (const call of allTrackedPowerQueryCalls(source)) {
    const scope = findEnclosingLetScope(rootScope, call.index);
    const visibleSteps = scope ? scope.visibleSteps : new Set();
    const args = splitTopLevelArgs(call.args);
    const sourceStepNames = call.functionName === 'Table.NestedJoin'
      ? [parsePowerQueryIdentifier(args[0]), parsePowerQueryIdentifier(args[2])].filter(Boolean)
      : [parsePowerQueryIdentifier(args[0])].filter(Boolean);
    for (const sourceStepName of sourceStepNames) {
      if (sourceStepName === '_') continue;
      if (!visibleSteps.has(sourceStepName.toLowerCase())) {
        diagnostics.push(`${call.functionName} references source step "${sourceStepName}", but that step is not defined in the query.`);
      }
    }
  }
  return diagnostics;
}

// Builds a tree of lexical let-scopes over [start, end) in `text` (absolute offsets
// throughout, so nested scopes never need re-slicing/re-indexing). Each scope carries its
// own step assignments plus `visibleSteps`, the union of its own steps and every ancestor
// scope's steps (matching M's nested-let visibility rules).
function buildLetScopeTree(text, start, end, ancestorSteps = [], parent = null) {
  const slice = text.slice(start, end);
  if (!/^\s*let\b/i.test(slice)) return null;
  const assignments = parsePowerQueryLetAssignments(slice);
  const inTarget = parsePowerQueryInTarget(slice);
  const ownSteps = new Set(assignments.map(step => step.name.toLowerCase()));
  const visibleSteps = new Set([...ancestorSteps, ...ownSteps]);
  const scope = { start, end, assignments, inTarget, ownSteps, visibleSteps, parent, nested: [] };
  const nestedRanges = findNestedLetRanges(text, start, end);
  scope.nested = nestedRanges
    .map(range => buildLetScopeTree(text, range.start, range.end, [...visibleSteps], scope))
    .filter(Boolean);
  return scope;
}

function walkLetScopeTree(scope, visit, isRoot = true) {
  visit(scope, isRoot);
  for (const child of scope.nested) walkLetScopeTree(child, visit, false);
}

// Finds the innermost scope whose [start, end) range contains `index`.
function findEnclosingLetScope(scope, index) {
  if (index < scope.start || index >= scope.end) return null;
  for (const child of scope.nested) {
    const found = findEnclosingLetScope(child, index);
    if (found) return found;
  }
  return scope;
}

// Finds `(let ... in ...)` expressions embedded as call arguments within [from, to) of
// `text`, returning the absolute [start, end) of each one's inner "let ... in ..." body
// (i.e. excluding the wrapping parens, so the range is itself a valid buildLetScopeTree input).
function findNestedLetRanges(text, from, to) {
  const ranges = [];
  const re = /\(\s*let\b/gi;
  re.lastIndex = from;
  let match;
  while ((match = re.exec(text)) && match.index < to) {
    const openParen = match.index;
    const closeParen = findMatchingParenIndex(text, openParen);
    if (closeParen < 0 || closeParen >= to) { re.lastIndex = match.index + 1; continue; }
    ranges.push({ start: openParen + 1, end: closeParen });
    re.lastIndex = closeParen + 1;
  }
  return ranges;
}

function findMatchingParenIndex(text, openIndex) {
  let depth = 0; let inString = false;
  for (let index = openIndex; index < text.length; index += 1) {
    const ch = text[index]; const next = text[index + 1];
    if (inString) { if (ch === '"' && next === '"') { index += 1; continue; } if (ch === '"') inString = false; continue; }
    if (ch === '"') { inString = true; continue; }
    else if (ch === '(') depth += 1;
    else if (ch === ')') { depth -= 1; if (depth === 0) return index; }
  }
  return -1;
}

function allTrackedPowerQueryCalls(source) {
  return [
    'Table.TransformColumnTypes', 'Table.SelectRows', 'Table.SelectColumns',
    'Table.RemoveColumns', 'Table.RenameColumns', 'Table.AddColumn',
    'Table.Sort', 'Table.Group', 'Table.NestedJoin',
  ].flatMap(functionName => findPowerQueryFunctionCalls(source, functionName).map(call => ({ ...call, functionName })))
    .sort((left, right) => left.index - right.index);
}

function parsePowerQueryLetAssignments(source) {
  const text = String(source || '');
  const letMatch = text.match(/^\s*let\b/i);
  if (!letMatch) return [];
  const inIndex = findTopLevelPowerQueryInIndex(text);
  if (inIndex < 0) return [];
  const chunks = splitTopLevelArgs(text.slice(letMatch[0].length, inIndex));
  const assignments = [];
  for (const chunk of chunks) {
    const eqIndex = findTopLevelEquals(chunk);
    if (eqIndex < 0) continue;
    const name = normalizePowerQueryIdentifier(chunk.slice(0, eqIndex).trim());
    if (name) assignments.push({ name, expression: chunk.slice(eqIndex + 1).trim() });
  }
  return assignments;
}

function parsePowerQueryInTarget(source) {
  const inIndex = findTopLevelPowerQueryInIndex(source);
  return inIndex < 0 ? null : normalizePowerQueryIdentifier(String(source).slice(inIndex + 2).trim());
}

function findTopLevelPowerQueryInIndex(source) {
  const text = String(source || '');
  let depthParen = 0; let depthBrace = 0; let depthBracket = 0; let inString = false; let candidate = -1;
  for (let index = 0; index < text.length; index += 1) {
    const ch = text[index]; const next = text[index + 1];
    if (inString) { if (ch === '"' && next === '"') index += 1; else if (ch === '"') inString = false; continue; }
    if (ch === '"') inString = true;
    else if (ch === '(') depthParen += 1; else if (ch === ')') depthParen -= 1;
    else if (ch === '{') depthBrace += 1; else if (ch === '}') depthBrace -= 1;
    else if (ch === '[') depthBracket += 1; else if (ch === ']') depthBracket -= 1;
    else if (depthParen === 0 && depthBrace === 0 && depthBracket === 0 && /\bin\b/i.test(text.slice(index, index + 2))) {
      const before = text[index - 1] || ' '; const after = text[index + 2] || ' ';
      if (!/[A-Za-z0-9_]/.test(before) && !/[A-Za-z0-9_]/.test(after)) candidate = index;
    }
  }
  return candidate;
}

function findTopLevelEquals(source) {
  const text = String(source || '');
  let depthParen = 0; let depthBrace = 0; let depthBracket = 0; let inString = false;
  for (let index = 0; index < text.length; index += 1) {
    const ch = text[index]; const next = text[index + 1];
    if (inString) { if (ch === '"' && next === '"') index += 1; else if (ch === '"') inString = false; continue; }
    if (ch === '"') inString = true;
    else if (ch === '(') depthParen += 1; else if (ch === ')') depthParen -= 1;
    else if (ch === '{') depthBrace += 1; else if (ch === '}') depthBrace -= 1;
    else if (ch === '[') depthBracket += 1; else if (ch === ']') depthBracket -= 1;
    else if (ch === '=' && depthParen === 0 && depthBrace === 0 && depthBracket === 0) return index;
  }
  return -1;
}

function findPowerQueryFunctionCalls(source, functionName) {
  const text = String(source || '');
  const calls = [];
  let searchFrom = 0;
  while (searchFrom < text.length) {
    const index = text.indexOf(functionName, searchFrom);
    if (index < 0) break;
    const open = text.indexOf('(', index + functionName.length);
    if (open < 0) break;
    let depth = 0; let inString = false;
    for (let cursor = open; cursor < text.length; cursor += 1) {
      const ch = text[cursor]; const next = text[cursor + 1];
      if (inString) { if (ch === '"' && next === '"') cursor += 1; else if (ch === '"') inString = false; continue; }
      if (ch === '"') inString = true;
      else if (ch === '(') depth += 1;
      else if (ch === ')') { depth -= 1; if (depth === 0) { calls.push({ index, args: text.slice(open + 1, cursor), preview: text.slice(index, Math.min(cursor + 1, index + 90)).replace(/\s+/g, ' ') }); searchFrom = cursor + 1; break; } }
    }
    if (searchFrom <= index) searchFrom = index + functionName.length;
  }
  return calls;
}

function splitTopLevelArgs(source) {
  const text = String(source || '');
  const args = [];
  let start = 0; let depthParen = 0; let depthBrace = 0; let depthBracket = 0; let inString = false;
  for (let index = 0; index < text.length; index += 1) {
    const ch = text[index]; const next = text[index + 1];
    if (inString) { if (ch === '"' && next === '"') index += 1; else if (ch === '"') inString = false; continue; }
    if (ch === '"') inString = true;
    else if (ch === '(') depthParen += 1; else if (ch === ')') depthParen -= 1;
    else if (ch === '{') depthBrace += 1; else if (ch === '}') depthBrace -= 1;
    else if (ch === '[') depthBracket += 1; else if (ch === ']') depthBracket -= 1;
    else if (ch === ',' && depthParen === 0 && depthBrace === 0 && depthBracket === 0) { args.push(text.slice(start, index).trim()); start = index + 1; }
  }
  args.push(text.slice(start).trim());
  return args;
}

function parsePowerQueryIdentifier(source) {
  return normalizePowerQueryIdentifier(String(source || '').trim());
}

function normalizePowerQueryIdentifier(source) {
  const raw = String(source || '').trim();
  if (!raw) return null;
  const quoted = raw.match(/^#?"((?:""|[^"])*)"$/);
  if (quoted) return quoted[1].replace(/""/g, '"');
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(raw) ? raw : null;
}

module.exports = {
  powerQueryStructuralDiagnostics,
  parsePowerQueryLetAssignments,
  parsePowerQueryInTarget,
  allTrackedPowerQueryCalls,
  parsePowerQueryIdentifier,
  normalizePowerQueryIdentifier,
  splitTopLevelArgs,
  findPowerQueryFunctionCalls,
  buildLetScopeTree,
  walkLetScopeTree,
  findEnclosingLetScope,
};
