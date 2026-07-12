const assert = require('node:assert/strict');
const test = require('node:test');
const { compileSqlToM } = require('../src/compiler');

const headerContext = {
  tables: {
    sales: {
      excelTableName: 'tblVentas',
      columns: [
        { name: 'status', header: 'Estado' },
        { name: 'amount', header: 'Monto Total', internalType: 'float' },
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
  assert.match(result.rejections[0].hint, /Simplify/);
});

test('accepts SELECT * when the report column set matches the table schema', () => {
  const result = compileSqlToM({
    sql: 'SELECT * FROM sales',
    report: { columns: [{ name: 'status' }, { name: 'amount' }] },
    headerContext,
  });
  assert.equal(result.ok, true, JSON.stringify(result.rejections));
  assert.match(result.mCode, /Table\.SelectColumns/);
});
