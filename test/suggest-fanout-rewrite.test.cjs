const assert = require('node:assert/strict');
const test = require('node:test');
const { compileSqlToM, suggestFanoutRewrite } = require('../src/compiler');

const headerContext = {
  tables: {
    sales: { columns: [{ name: 'id', isPrimaryKey: true }, { name: 'status' }] },
    children: { columns: [{ name: 'id' }, { name: 'sale_id' }, { name: 'amount' }] },
    payments: { columns: [{ name: 'id' }, { name: 'sale_id' }, { name: 'amount' }] },
  },
};

function assertRecompiles(sql, columns) {
  const suggestion = suggestFanoutRewrite(sql);
  assert.equal(suggestion.ok, true, `expected a suggestion, got bail: ${suggestion.reason}`);
  const result = compileSqlToM({ sql: suggestion.sql, report: { columns: columns.map(name => ({ name })) }, headerContext });
  assert.equal(result.ok, true, `suggested SQL failed to recompile: ${JSON.stringify(result.rejections)}\n${suggestion.sql}`);
  return suggestion;
}

test('suggests a Pattern-D rewrite for the canonical two-branch join fan-out', () => {
  const sql = `SELECT s.status, SUM(c.amount) AS child_total, SUM(p.amount) AS payment_total
FROM sales s
LEFT JOIN children c ON s.id = c.sale_id
LEFT JOIN payments p ON s.id = p.sale_id
GROUP BY s.status`;
  assertRecompiles(sql, ['status', 'child_total', 'payment_total']);
});

test('rewrites a COUNT over a branch as an outer SUM of per-key counts', () => {
  const sql = `SELECT s.status, COUNT(c.id) AS child_count, SUM(p.amount) AS payment_total
FROM sales s
LEFT JOIN children c ON s.id = c.sale_id
LEFT JOIN payments p ON s.id = p.sale_id
GROUP BY s.status`;
  const suggestion = assertRecompiles(sql, ['status', 'child_count', 'payment_total']);
  assert.match(suggestion.sql, /SUM\(t\.child_count\)/);
});

test('bails on AVG over a joined branch instead of guessing', () => {
  const sql = `SELECT s.status, AVG(c.amount) AS avg_amount, SUM(p.amount) AS payment_total
FROM sales s
LEFT JOIN children c ON s.id = c.sale_id
LEFT JOIN payments p ON s.id = p.sale_id
GROUP BY s.status`;
  const result = suggestFanoutRewrite(sql);
  assert.equal(result.ok, false);
  assert.match(result.reason, /AVG/);
});

test('bails on an aggregate mixing two joined branches', () => {
  const sql = `SELECT s.status, SUM(c.amount * p.amount) AS mixed
FROM sales s
LEFT JOIN children c ON s.id = c.sale_id
LEFT JOIN payments p ON s.id = p.sale_id
GROUP BY s.status`;
  const result = suggestFanoutRewrite(sql);
  assert.equal(result.ok, false);
  assert.match(result.reason, /mixes columns/);
});

test('bails on unqualified column references', () => {
  const sql = `SELECT status, SUM(c.amount) AS child_total, SUM(p.amount) AS payment_total
FROM sales s
LEFT JOIN children c ON s.id = c.sale_id
LEFT JOIN payments p ON s.id = p.sale_id
GROUP BY status`;
  const result = suggestFanoutRewrite(sql);
  assert.equal(result.ok, false);
  assert.match(result.reason, /unqualified/);
});
