const assert = require('node:assert/strict');
const test = require('node:test');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const CLI = path.join(__dirname, '../skills/sql-to-power-m/scripts/compile.cjs');
function run(args) {
  const result = spawnSync(process.execPath, [CLI, ...args], { encoding: 'utf8' });
  return result;
}

test('CLI compiles SQL against an inline schema', () => {
  const result = run([
    '--sql', 'SELECT status, SUM(amount) AS total FROM sales GROUP BY status',
    '--schema', JSON.stringify({ tables: { sales: { excelTableName: 'tblVentas', columns: [{ name: 'status' }, { name: 'amount', internalType: 'float' }] } } }),
  ]);
  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.ok, true);
  assert.match(output.mCode, /Excel\.CurrentWorkbook\(\)\{\[Name="tblVentas"\]\}/);
  assert.equal(output.suggestion, null);
});

test('CLI surfaces a suggestion alongside a join fan-out rejection', () => {
  const sql = `SELECT s.status, SUM(c.amount) AS child_total, SUM(p.amount) AS payment_total
FROM sales s
LEFT JOIN children c ON s.id = c.sale_id
LEFT JOIN payments p ON s.id = p.sale_id
GROUP BY s.status`;
  const schema = { tables: {
    sales: { columns: [{ name: 'id' }, { name: 'status' }] },
    children: { columns: [{ name: 'id' }, { name: 'sale_id' }, { name: 'amount' }] },
    payments: { columns: [{ name: 'id' }, { name: 'sale_id' }, { name: 'amount' }] },
  } };
  const result = run(['--sql', sql, '--schema', JSON.stringify(schema)]);
  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.ok, false);
  assert.equal(output.rejections[0].construct, 'join fan-out');
  assert.equal(output.suggestion.ok, true);
  assert.match(output.suggestion.sql, /GROUP BY t\.status/);
});

test('CLI runs --simplify-only without any schema at all', () => {
  const sql = `SELECT s.status, SUM(c.amount) AS child_total, SUM(p.amount) AS payment_total
FROM sales s
LEFT JOIN children c ON s.id = c.sale_id
LEFT JOIN payments p ON s.id = p.sale_id
GROUP BY s.status`;
  const result = run(['--simplify-only', '--sql', sql]);
  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.suggestion.ok, true);
  assert.match(output.suggestion.sql, /FROM \(/);
});

test('CLI exits non-zero with a clear message when no schema is given and --simplify-only is not set', () => {
  const result = run(['--sql', 'SELECT 1']);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /No schema provided/);
});

test('CLI exits non-zero with a clear message when no SQL is given', () => {
  const result = run(['--schema', '{"tables":{}}']);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /No SQL provided/);
});
