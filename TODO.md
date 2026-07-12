# TODO

## MVP follow-up

- [x] Browser-only SQL editor and SQL→M compile flow
- [x] Upload `.xlsx`/`.xlsm` and detect worksheet headers
- [x] Keep the Excel Table name editable and visible
- [x] Manual schema JSON import/export path
- [x] Query starter for a table, selected columns, and a simple filter
- [x] Copy M and download `.pq`
- [x] Structured compiler rejection display
- [x] GitHub Pages build/deploy workflow
- [x] Compiler smoke tests and schema extraction tests

## Next product increments

- [ ] Add a bundled/offline XLSX reader or a small ZIP/XML table-metadata reader so the app does not rely on the SheetJS CDN.
- [ ] Improve `.xlsx` table metadata extraction and show a warning when a sheet has multiple Excel Tables.
- [ ] Add a visual builder for supported joins, grouping, aggregates, `CASE`, and `ORDER BY`.
- [ ] Add schema aliases and lookup metadata editing for SQLForge-style collapsed lookups.
- [ ] Add a supported-SQL grammar/reference page generated from the parser and validator.
- [ ] Add shareable, URL-encoded SQL + schema links without storing workbook contents.
- [ ] Add a browser test using Playwright for upload → compile → copy/download flows.
- [ ] Add a compiler version badge and a documented sync process from the parent `IS` project.

## Deliberately deferred

- [ ] Cloudflare Worker for compilation — not needed while parsing and compilation are small and local.
- [ ] Browser execution of Power Query M — Excel remains the runtime of record.
- [ ] Uploading workbook rows to a server — out of scope for privacy and unnecessary for schema extraction.
