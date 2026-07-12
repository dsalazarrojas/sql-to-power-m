const SQL_RESERVED_WORDS = new Set([
  'select', 'from', 'where', 'group', 'by', 'having', 'order', 'join', 'left', 'inner',
  'right', 'full', 'outer', 'cross', 'on', 'union', 'all', 'and', 'or', 'not', 'null',
]);

export function toSqlName(value, fallback = 'column') {
  const clean = String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9_$]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toLowerCase();
  const result = clean || fallback;
  return /^\d/.test(result) || SQL_RESERVED_WORDS.has(result) ? `_${result}` : result;
}

export function uniqueName(value, used, fallback) {
  const base = toSqlName(value, fallback);
  let candidate = base;
  let suffix = 2;
  while (used.has(candidate.toLowerCase())) candidate = `${base}_${suffix++}`;
  used.add(candidate.toLowerCase());
  return candidate;
}

export function inferInternalType(values) {
  const present = values.filter(value => value !== null && value !== undefined && String(value).trim() !== '');
  if (!present.length) return 'string';
  if (present.every(value => value instanceof Date || (typeof value === 'string' && /^\d{4}-\d{1,2}-\d{1,2}(?:[T ]|$)/.test(value)))) return 'date';
  if (present.every(value => typeof value === 'number' || (typeof value === 'string' && /^[-+]?\d+(?:\.\d+)?$/.test(value.trim())))) return 'float';
  return 'string';
}

function headerText(value, index) {
  const text = String(value ?? '').trim();
  return text || `Column ${index + 1}`;
}

export function sheetToTableSchema(sheet, sheetName, excelTableName, xlsx) {
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
  const headerIndex = rows.findIndex(row => Array.isArray(row) && row.some(value => value !== null && value !== undefined && String(value).trim() !== ''));
  const headers = headerIndex < 0 ? [] : rows[headerIndex];
  const dataRows = headerIndex < 0 ? [] : rows.slice(headerIndex + 1);
  const used = new Set();
  const columns = headers.map((value, index) => {
    const header = headerText(value, index);
    const name = uniqueName(header, used, `column_${index + 1}`);
    return { name, header, internalType: inferInternalType(dataRows.map(row => row?.[index])) };
  });
  return {
    name: toSqlName(sheetName, 'table'),
    excelTableName: String(excelTableName || sheetName || 'Table1').trim() || 'Table1',
    columns,
    rowCount: dataRows.length,
  };
}

function normalizeZipPath(value) {
  const parts = [];
  for (const part of String(value).replaceAll('\\', '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop(); else parts.push(part);
  }
  return parts.join('/');
}

function zipText(workbook, path) {
  const files = workbook?.files;
  if (!files) return null;
  const key = normalizeZipPath(path);
  const file = files instanceof Map ? files.get(key) : files[key];
  if (!file) return null;
  const content = file.content ?? file;
  if (typeof content === 'string') return content;
  if (content instanceof Uint8Array || ArrayBuffer.isView(content)) return new TextDecoder().decode(content);
  return null;
}

function attr(xml, name) {
  const match = String(xml).match(new RegExp(`${name}\\s*=\\s*["']([^"']+)["']`, 'i'));
  return match?.[1] || null;
}

function relationshipMap(xml) {
  const map = new Map();
  for (const match of String(xml || '').matchAll(/<Relationship\b[^>]*>/gi)) {
    const tag = match[0];
    const id = attr(tag, 'Id');
    const target = attr(tag, 'Target');
    if (id && target) map.set(id, target);
  }
  return map;
}

function extractTableNames(workbook) {
  const names = new Map();
  const workbookXml = zipText(workbook, 'xl/workbook.xml');
  const workbookRels = relationshipMap(zipText(workbook, 'xl/_rels/workbook.xml.rels'));
  if (!workbookXml || !workbookRels.size) return names;
  for (const sheetMatch of workbookXml.matchAll(/<sheet\b[^>]*>/gi)) {
    const sheetTag = sheetMatch[0];
    const sheetName = attr(sheetTag, 'name');
    const relationshipId = attr(sheetTag, 'r:id');
    const sheetTarget = workbookRels.get(relationshipId);
    if (!sheetName || !sheetTarget) continue;
    const sheetPath = normalizeZipPath(`xl/${sheetTarget.replace(/^\.\.\//, '')}`);
    const sheetXml = zipText(workbook, sheetPath);
    const relPath = normalizeZipPath(`${sheetPath.slice(0, sheetPath.lastIndexOf('/') + 1)}_rels/${sheetPath.slice(sheetPath.lastIndexOf('/') + 1)}.rels`);
    const sheetRels = relationshipMap(zipText(workbook, relPath));
    const tablePart = sheetXml?.match(/<tablePart\b[^>]*>/i);
    const tableRelId = tablePart && attr(tablePart[0], 'r:id');
    const tableTarget = tableRelId && sheetRels.get(tableRelId);
    if (!tableTarget) continue;
    const tablePath = normalizeZipPath(`${sheetPath.slice(0, sheetPath.lastIndexOf('/') + 1)}${tableTarget}`);
    const tableXml = zipText(workbook, tablePath);
    const tableName = tableXml && (attr(tableXml.match(/<table\b[^>]*>/i)?.[0], 'displayName') || attr(tableXml.match(/<table\b[^>]*>/i)?.[0], 'name'));
    if (tableName) names.set(sheetName, tableName);
  }
  return names;
}

export function workbookToHeaderContext(workbook, xlsx) {
  const tables = {};
  const warnings = [];
  const actualNames = extractTableNames(workbook);
  const hasActualTables = actualNames.size > 0;
  let skippedWorksheets = 0;
  const usedTableNames = new Set();
  for (const sheetName of workbook.SheetNames || []) {
    if (hasActualTables && !actualNames.has(sheetName)) {
      skippedWorksheets += 1;
      continue;
    }
    const sheet = workbook.Sheets?.[sheetName];
    if (!sheet) continue;
    const schema = sheetToTableSchema(sheet, sheetName, actualNames.get(sheetName) || sheetName, xlsx);
    if (!schema.columns.length) continue;
    let key = schema.name;
    let suffix = 2;
    while (tables[key]) key = `${schema.name}_${suffix++}`;
    if (key !== schema.name) warnings.push(`The worksheet “${sheetName}” was renamed to SQL table “${key}” to avoid a duplicate.`);
    if (usedTableNames.has(schema.excelTableName.toLowerCase())) warnings.push(`The Excel Table name “${schema.excelTableName}” appears more than once; update the mapping before compiling.`);
    usedTableNames.add(schema.excelTableName.toLowerCase());
    tables[key] = { excelTableName: schema.excelTableName, columns: schema.columns };
  }
  if (hasActualTables && skippedWorksheets) warnings.push(`${skippedWorksheets} worksheet(s) without Excel Table metadata were omitted.`);
  if (!hasActualTables && Object.keys(tables).length) warnings.push('No Excel Table metadata was found; worksheet names were used as Excel Table names. Confirm these names in Excel Table Design.');
  if (!Object.keys(tables).length) warnings.push('No non-empty worksheet headers were found. Add a header row or use the manual schema JSON editor.');
  return { headerContext: { tables }, warnings };
}
