import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// dist/Code.gs は export を除去した Apps Script 形式なので、
// テストでは rowsToObjects の定義だけを抜き出して評価する。
const code = readFileSync(new URL('../dist/Code.gs', import.meta.url), 'utf8');
const m = code.match(/function rowsToObjects[\s\S]*?\n\}/);
assert.ok(m, 'rowsToObjects not found in dist/Code.gs');
const rowsToObjects = new Function(`${m[0]}; return rowsToObjects;`)();

test('ヘッダー行をキーにしてオブジェクト化する', () => {
  const values = [
    ['id', 'zh', 'ja'],
    ['c001', '你好', 'こんにちは'],
    ['c002', '谢谢', 'ありがとう'],
  ];
  assert.deepEqual(rowsToObjects(values), [
    { id: 'c001', zh: '你好', ja: 'こんにちは' },
    { id: 'c002', zh: '谢谢', ja: 'ありがとう' },
  ]);
});

test('空行はスキップする', () => {
  const values = [
    ['id', 'zh'],
    ['', ''],
    ['c001', '你好'],
  ];
  assert.deepEqual(rowsToObjects(values), [{ id: 'c001', zh: '你好' }]);
});

test('データが無い場合は空配列', () => {
  assert.deepEqual(rowsToObjects([]), []);
  assert.deepEqual(rowsToObjects([['id']]), []);
  assert.deepEqual(rowsToObjects(null), []);
});
