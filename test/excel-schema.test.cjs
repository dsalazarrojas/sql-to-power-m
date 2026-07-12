const assert = require('node:assert/strict');
const test = require('node:test');

test('normalizes names and avoids collisions', async () => {
  const { toSqlName, uniqueName, inferInternalType } = await import('../src/excel-schema.js');
  assert.equal(toSqlName('Monto Total'), 'monto_total');
  assert.equal(toSqlName('SELECT'), '_select');
  const used = new Set();
  assert.equal(uniqueName('Order ID', used, 'column'), 'order_id');
  assert.equal(uniqueName('Order ID', used, 'column'), 'order_id_2');
  assert.equal(inferInternalType([1, 2, 3]), 'float');
  assert.equal(inferInternalType(['2026-07-12', '2026-07-13']), 'date');
  assert.equal(inferInternalType(['paid', 'open']), 'string');
});

test('builds a schema from worksheet rows without retaining workbook rows', async () => {
  const { sheetToTableSchema } = await import('../src/excel-schema.js');
  const sheet = { rows: [['Order ID', 'Amount'], ['A-1', 12], ['A-2', 18]] };
  const fakeXlsx = { utils: { sheet_to_json: value => value.rows } };
  const schema = sheetToTableSchema(sheet, 'Orders', 'tblOrders', fakeXlsx);
  assert.deepEqual(schema.columns.map(column => column.name), ['order_id', 'amount']);
  assert.deepEqual(schema.columns.map(column => column.header), ['Order ID', 'Amount']);
  assert.equal(schema.columns[1].internalType, 'float');
  assert.equal(schema.excelTableName, 'tblOrders');
  assert.equal('rows' in schema, false);
});
