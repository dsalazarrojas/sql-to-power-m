const assert = require('node:assert/strict');
const test = require('node:test');
const { compileSqlToM } = require('../src/compiler');

const headerContext = {
  tables: {
    sales: {
      excelTableName: 'tblVentas',
      columns: [
        { name: 'id', header: 'Id', internalType: 'integer' },
        { name: 'status', header: 'Estado' },
        { name: 'amount', header: 'Monto Total', internalType: 'float' },
        { name: 'created_date', header: 'Fecha', internalType: 'string' },
        { name: 'closed_date', header: 'Fecha cierre', internalType: 'string' },
      ],
    },
    payments: {
      excelTableName: 'tblPagos',
      columns: [
        { name: 'sale_id', header: 'Venta Id', internalType: 'integer' },
        { name: 'amount', header: 'Monto', internalType: 'float' },
      ],
    },
  },
};

test('compiles a grouped query against an Excel Table', () => {
  const result = compileSqlToM({
    sql: 'SELECT status, SUM(amount) AS total FROM sales GROUP BY status ORDER BY total DESC',
    report: { columns: [] },
    headerContext,
  });
  assert.equal(result.ok, true, JSON.stringify(result.rejections));
  assert.match(result.mCode, /Excel\.CurrentWorkbook\(\)\{\[Name="tblVentas"\]\}/);
  assert.match(result.mCode, /Table\.Group/);
  assert.match(result.mCode, /Order\.Descending/);
});

test('returns a useful structured rejection for unsupported SQL', () => {
  const result = compileSqlToM({
    sql: 'SELECT * FROM sales LIMIT 10',
    report: { columns: [] },
    headerContext,
  });
  assert.equal(result.ok, false);
  assert.equal(result.rejections[0].construct, 'LIMIT');
  assert.match(result.rejections[0].hint, /LIMIT\/OFFSET/);
});

test('accepts SELECT * when the report column set matches the table schema', () => {
  const result = compileSqlToM({
    sql: 'SELECT * FROM sales',
    report: { columns: [{ name: 'id' }, { name: 'status' }, { name: 'amount' }, { name: 'created_date' }, { name: 'closed_date' }] },
    headerContext,
  });
  assert.equal(result.ok, true, JSON.stringify(result.rejections));
  assert.match(result.mCode, /Table\.SelectColumns/);
});

test('aggregates an arithmetic expression directly, without a staging subquery', () => {
  const result = compileSqlToM({
    sql: "SELECT status, ROUND(AVG(julianday(closed_date) - julianday(created_date)), 1) AS avg_days FROM sales WHERE closed_date IS NOT NULL GROUP BY status",
    report: { columns: [] },
    headerContext,
  });
  assert.equal(result.ok, true, JSON.stringify(result.rejections));
  assert.match(result.mCode, /List\.Transform\(Table\.ToRecords\(_\)/);
});

test('compiles a multi-condition JOIN ... ON clause', () => {
  const result = compileSqlToM({
    sql: "SELECT s.status, p.amount FROM sales s LEFT JOIN payments p ON p.sale_id = s.id AND p.amount > 0",
    report: { columns: [] },
    headerContext,
  });
  assert.equal(result.ok, true, JSON.stringify(result.rejections));
  assert.doesNotMatch(result.mCode, /undefined/);
});

test('rewrites NOT EXISTS into an anti-join', () => {
  const result = compileSqlToM({
    sql: "SELECT s.status FROM sales s WHERE NOT EXISTS (SELECT 1 FROM payments p WHERE p.sale_id = s.id)",
    report: { columns: [] },
    headerContext,
  });
  assert.equal(result.ok, true, JSON.stringify(result.rejections));
  assert.match(result.mCode, /JoinKind\.LeftOuter/);
  assert.match(result.mCode, /= null/);
});
