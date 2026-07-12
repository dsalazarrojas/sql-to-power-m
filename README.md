# SQL → Power Query M

A privacy-first GitHub Pages tool for turning the supported SQL subset from SQLForge into Power Query M that reads Excel Tables through `Excel.CurrentWorkbook()`.

## Product direction

The first release is deliberately a focused compiler playground rather than a full SQL database or an Excel runtime:

1. The user uploads an `.xlsx`/`.xlsm` file, or enters schema JSON manually.
2. The browser extracts worksheet headers and infers lightweight types. It never uploads workbook data.
3. The user reviews the SQL names and real Excel Table names, writes SQL, or generates a simple query starter.
4. The browser runs the same tested SQL→M compiler used by the parent SQLForge project.
5. The user copies or downloads the generated `.pq` query and pastes it into Excel Power Query.

The compiler snapshot lives in [`src/compiler/`](src/compiler/). It is copied from `../IS/lib/generation/sql-to-m/` and `../IS/lib/generation/power-query-lint.js`; the app does not depend on the SQLForge server, a database, a Cloudflare Worker, or a secret.

`SELECT *` is accepted when it is a direct query against a detected table, because the app can pass the exact detected column set to the compiler. For joins or more complex shapes, prefer an explicit projection so the resulting M columns stay predictable.

## Run locally

```bash
npm install
npm run build
python3 -m http.server 4173 -d dist
```

Open <http://localhost:4173>. The `.xlsx` reader is loaded from the SheetJS browser CDN at runtime. If that CDN is unavailable, the manual schema editor and sample schema still work.

## GitHub Pages

The workflow in [`.github/workflows/pages.yml`](.github/workflows/pages.yml) builds and deploys `dist/` on pushes to `main`. The site is live at **[https://gic.mx/sql-to-power-m/](https://gic.mx/sql-to-power-m/)**.

## Important boundaries

This is a SQL subset for Power Query, not a general SQL engine. The compiler intentionally rejects constructs such as CTEs, `LIMIT`, `RIGHT JOIN`, `LIKE`, window functions, `IN (SELECT ...)`, self-joins, and unsafe join fan-out aggregates. A successful compile produces M source; it does not refresh Excel or execute M in the browser.

An uploaded worksheet is initially treated as an Excel Table whose name matches the worksheet name. For reliable output, replace that value with the actual Excel Table name shown in Excel's **Table Design → Table Name**. The schema panel makes this mapping explicit so a pasted query does not silently reference the wrong object.

## Design improvements over the original idea

- Browser-first privacy: no workbook bytes leave the device, and a Worker is deferred until a concrete need exists.
- Schema-first workflow: the tool makes the required SQL-name → display-header → Excel-Table mapping visible instead of hiding it behind a parser error.
- Progressive disclosure: SQL editor first, query starter for newcomers, schema JSON escape hatch for advanced users.
- Honest compiler UX: structured rejection messages include the unsupported construct, hint, and source position.
- Deployable artifact: the Pages site is static; Webpack only bundles the compiler and UI JavaScript.
