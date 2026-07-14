#!/usr/bin/env node
'use strict';
// CLI wrapper around the same browser compiler that runs on gic.mx/sql-to-power-m/.
// Reads SQL (+ optionally a schema/headerContext), prints one JSON result line to stdout.
// Never touches a network or an Excel file - it is a pure text-in, text-out compiler.
const fs = require('fs');
const path = require('path');
const { compileSqlToM, suggestFanoutRewrite } = require('../../../src/compiler');

function parseArgs(argv) {
  const args = { columns: null };
  for (let i = 0; i < argv.length; i++) {
    const flag = argv[i];
    const value = () => argv[++i];
    if (flag === '--sql') args.sql = value();
    else if (flag === '--sql-file') args.sqlFile = value();
    else if (flag === '--schema') args.schema = value();
    else if (flag === '--schema-file') args.schemaFile = value();
    else if (flag === '--columns') args.columns = value().split(',').map(s => s.trim()).filter(Boolean);
    else if (flag === '--simplify-only') args.simplifyOnly = true;
    else if (flag === '--help' || flag === '-h') args.help = true;
    else { process.stderr.write(`Unknown argument: ${flag}\n`); process.exit(2); }
  }
  return args;
}

function readStdinSync() {
  try { return fs.readFileSync(0, 'utf8'); } catch { return ''; }
}

const USAGE = `Usage: node compile.cjs [options]

  --sql <text>          Inline SQL. If omitted (and no --sql-file), reads SQL from stdin.
  --sql-file <path>      Path to a .sql file.
  --schema <json>        Inline JSON headerContext, e.g. {"tables":{"sales":{"columns":[...]}}}.
  --schema-file <path>   Path to a JSON file with the same shape.
  --columns <a,b,c>      Optional report.columns list (only needed for a bare SELECT * query).
  --simplify-only        Skip compilation; only run the join-fan-out rewrite suggestion.
                         Use this when you don't have Excel Table headers on hand yet -
                         it works from the SQL text alone.

Prints one line of JSON to stdout:
  { ok, mCode, rejections, suggestion }

suggestion is present whenever a "join fan-out" rejection occurs (or always,
under --simplify-only): { ok: true, sql, notes } or { ok: false, reason }.
A suggestion is text to review and paste back in, never something this tool
adopts on its own - see SKILL.md for how to use it in an agent workflow.
`;

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) { process.stdout.write(USAGE); return; }

  let sql = args.sql;
  if (!sql && args.sqlFile) sql = fs.readFileSync(path.resolve(args.sqlFile), 'utf8');
  if (!sql) sql = readStdinSync();
  sql = String(sql || '').trim();
  if (!sql) { process.stderr.write('No SQL provided (use --sql, --sql-file, or stdin).\n\n' + USAGE); process.exit(2); }

  if (args.simplifyOnly) {
    process.stdout.write(JSON.stringify({ ok: null, mCode: null, rejections: null, suggestion: suggestFanoutRewrite(sql) }) + '\n');
    return;
  }

  let headerContext = null;
  if (args.schema) headerContext = JSON.parse(args.schema);
  else if (args.schemaFile) headerContext = JSON.parse(fs.readFileSync(path.resolve(args.schemaFile), 'utf8'));
  if (!headerContext) {
    process.stderr.write('No schema provided. Either pass --schema/--schema-file, or pass --simplify-only to just get a join-fan-out rewrite suggestion from the SQL text alone.\n\n' + USAGE);
    process.exit(2);
  }

  const report = { columns: (args.columns || []).map(name => ({ name })) };
  const result = compileSqlToM({ sql, report, headerContext });
  if (!result.ok && result.rejections.some(r => r.construct === 'join fan-out')) {
    result.suggestion = suggestFanoutRewrite(sql);
  } else {
    result.suggestion = null;
  }
  process.stdout.write(JSON.stringify(result) + '\n');
}

main();
