'use strict';

// Deliberately small SQL lexer. Positions are zero-based character offsets.
function tokenize(sql) {
  const text = String(sql || ''); const tokens = []; let i = 0;
  const add = (type, value, position) => tokens.push({ type, value, position });
  while (i < text.length) {
    if (/\s/.test(text[i])) { i += 1; continue; }
    const p = i;
    if (text.slice(i, i + 2) === '--') { i = text.indexOf('\n', i + 2); if (i < 0) break; continue; }
    if (text[i] === "'" || text[i] === '"') {
      const quote = text[i++]; let value = '';
      while (i < text.length) { if (text[i] === quote && text[i + 1] === quote) { value += quote; i += 2; } else if (text[i] === quote) { i += 1; break; } else value += text[i++]; }
      add(quote === "'" ? 'string' : 'identifier', value, p); continue;
    }
    const number = text.slice(i).match(/^\d+(?:\.\d+)?/); if (number) { add('number', number[0], p); i += number[0].length; continue; }
    const word = text.slice(i).match(/^[A-Za-z_][A-Za-z0-9_$]*/); if (word) { add('word', word[0], p); i += word[0].length; continue; }
    const op = ['<=', '>=', '<>', '!=', '||'].find(x => text.slice(i, i + x.length) === x);
    if (op) { add('operator', op, p); i += op.length; continue; }
    if ('(),.*+-/=<>;'.includes(text[i])) { add('symbol', text[i++], p); continue; }
    add('unknown', text[i++], p);
  }
  tokens.push({ type: 'eof', value: '', position: text.length }); return tokens;
}
module.exports = { tokenize };
