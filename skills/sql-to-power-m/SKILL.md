---
name: sql-to-power-m
description: >
  Converts a practical SQL subset (SELECT/JOIN/WHERE/GROUP BY/HAVING/ORDER
  BY/CASE/UNION over Excel Tables) into Power Query M for
  Excel.CurrentWorkbook(), and detects and reshapes "join fan-out" SQL - a
  query that aggregates columns from two or more joined 1:N child tables and
  would double-count rows - into a safe pre-aggregate/join/wrap form. Trigger
  when the user asks to "convert this SQL to Power Query", "turn this into a
  Power Query M query for Excel", pastes a SQL query and a "join fan-out" or
  similar Power Query aggregation error, or asks to "simplify"/"fix"/"reshape"
  a SQL query that joins multiple child tables and sums/counts across them.
  Same compiler that powers https://gic.mx/sql-to-power-m/.
---

# SKILL — SQL to Power Query M

This skill wraps the same browser-only compiler behind
[gic.mx/sql-to-power-m](https://gic.mx/sql-to-power-m/) (source in
`src/compiler/` in this repo) as a CLI, so it can be driven from an agent
session without a browser.

It does two related things:

1. **Compile** a supported SQL subset into Power Query M that reads Excel
   Tables via `Excel.CurrentWorkbook()`.
2. **Simplify** a SQL query that hits the `join fan-out` rejection - aggregating
   columns from two or more joined 1:N child tables over the same parent row,
   which would double-count through a naive join - into the equivalent safe
   shape (pre-aggregate each branch, join those, wrap in an outer query).

Both are deterministic, offline, text-in/text-out operations. There is no
network call and no Excel automation involved anywhere in this skill.

## When to use this

- The user has a SQL query (or is willing to write one) and wants Power Query
  M for an Excel workbook.
- The user pastes a query and a Power Query/Excel error about a report that
  double-counts, or a "join fan-out" compiler rejection, and wants it fixed.
- The user asks you to simplify or reshape a query that joins multiple child
  tables and aggregates across more than one of them.

## How to run it

```bash
node scripts/compile.cjs --sql "<SQL>" --schema '<headerContext JSON>'
# or, from files:
node scripts/compile.cjs --sql-file query.sql --schema-file schema.json
# or, when you don't have Excel Table headers yet and just want the reshape:
node scripts/compile.cjs --simplify-only --sql "<SQL>"
```

Run `node scripts/compile.cjs --help` for the full flag reference. The script
prints one line of JSON to stdout:

```json
{ "ok": true, "mCode": "let ... in ...", "rejections": [], "suggestion": null }
```

or, on rejection:

```json
{ "ok": false, "mCode": null, "rejections": [{ "construct": "...", "message": "...", "hint": "..." }], "suggestion": null }
```

`suggestion` is populated whenever the rejection's `construct` is
`"join fan-out"` (or always, under `--simplify-only`):

```json
{ "ok": true, "sql": "SELECT t.status ... FROM ( ... ) t GROUP BY t.status", "notes": ["..."] }
```

or, when the shape is too complex to reshape safely:

```json
{ "ok": false, "reason": "AVG over joined branch \"c\" cannot be safely decomposed by auto-suggestion; compute SUM and COUNT separately and divide by hand." }
```

## The schema (headerContext)

The compiler needs to know the SQL table/column names and, for a full
compile, the real Excel Table name and display header for each column (the
`excelTableName`/`header` fields) - these become the `Excel.CurrentWorkbook()`
table reference and the M step names. Minimal shape:

```json
{
  "tables": {
    "sales": {
      "excelTableName": "tblSales",
      "columns": [
        { "name": "id", "header": "Id", "internalType": "integer", "isPrimaryKey": true },
        { "name": "status", "header": "Status", "internalType": "string" }
      ]
    }
  }
}
```

If the user hasn't given you this yet, ask for it, or point them at the
"Upload workbook" flow on the website (it reads worksheet headers locally in
the browser and never uploads workbook bytes). If you only need a
**simplification suggestion** and don't have the schema, use
`--simplify-only` - the rewrite is derived purely from the SQL's own
structure (aliases, join conditions, aggregate expressions), not from a
schema.

## What "simplify" means here, precisely

The `join fan-out` shape is: 2+ joined tables where some `SUM`/`COUNT`/`AVG`/
`MIN`/`MAX` reads a column from a joined (non-`FROM`) alias. Naively joining
two 1:N children onto one parent multiplies rows through a cross product, so
summing either child's measure double- or triple-counts it. The fix -
documented in `docs/excel-report-sql-patterns.md` in the parent SQLForge
project, and what this tool automates - is to pre-aggregate each branch to
one row per key, join those pre-aggregated branches instead of the raw child
tables, then wrap the whole thing in an outer query that does the final
grouping.

**This tool never guesses.** When the shape is a genuine limitation rather
than something it knows how to reshape (an aggregate mixing columns from two
different branches, `AVG` over a branch - which can't be composed by simply
re-averaging an average, a join keyed on a compound or cross-alias condition,
a column used both inside and outside its own aggregate, a `WHERE` filter on
a joined branch column), it returns `{ "ok": false, "reason": "..." }` with a
plain-language explanation instead of emitting SQL it isn't sure is correct.
**Relay that reason to the user verbatim rather than attempting the reshape
yourself** - if the tool declined, treat it as a real signal that the
shape needs a human decision, not a nudge to route around it.

## What to do with a suggestion

A suggestion is SQL text, not something this skill (or you) should adopt
silently:

1. Show the user the suggested SQL and explain what changed (pre-aggregate
   each branch, then join and wrap).
2. Confirm the columns/aliases still make sense for their schema.
3. Re-run the compile step on the suggested SQL (with the real schema, not
   `--simplify-only`) to confirm it now compiles - `ok: true` is the real
   acceptance test, not the suggestion step itself.
4. Only then hand back the `mCode` for them to paste into Power Query.

## What this skill does not do

- It does not execute M or refresh Excel - it only emits M source text.
- It does not read `.xlsx`/`.xlsm` files - that's the website's job (SheetJS
  in the browser). If the user has a workbook and no schema JSON yet, point
  them at the "Upload workbook" flow at gic.mx/sql-to-power-m/, which never
  uploads the file anywhere.
- It rejects (does not attempt) CTEs, `RIGHT JOIN`, `LIMIT`/`OFFSET`, `LIKE`/
  `GLOB`, window functions, `IN (SELECT ...)`, and self-joins - these are
  hard rejects in the compiler, not things `--simplify-only` can fix.
