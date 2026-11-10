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

test('列が不足する行は空文字で埋める', () => {
  const values = [
    ['id', 'zh', 'ja'],
    ['c001', '你好'],
  ];
  assert.deepEqual(rowsToObjects(values), [{ id: 'c001', zh: '你好', ja: '' }]);
});

test('数値セルはそのまま保持する', () => {
  const values = [
    ['id', 'count'],
    ['c001', 3],
  ];
  assert.deepEqual(rowsToObjects(values), [{ id: 'c001', count: 3 }]);
});

test('null セルが混ざる行は空行扱いしない', () => {
  const values = [
    ['id', 'zh'],
    [null, '你好'],
  ];
  assert.deepEqual(rowsToObjects(values), [{ id: null, zh: '你好' }]);
});

function extractFn(name) {
  const re = new RegExp(`function ${name}[\\s\\S]*?\\n\\}`);
  const mm = code.match(re);
  assert.ok(mm, `${name} not found in dist/Code.gs`);
  return new Function(`${mm[0]}; return ${name};`)();
}

test('buildJudgePrompt は質問・模範解答・回答とJSON指定を含む', () => {
  const buildJudgePrompt = extractFn('buildJudgePrompt');
  const p = buildJudgePrompt({
    question: '你昨天工作了几个小时？',
    qword: '几',
    expected: '我昨天工作了八个小时。',
    answer: '我昨天工作八个小时。',
  });
  assert.ok(p.includes('你昨天工作了几个小时？'));
  assert.ok(p.includes('我昨天工作了八个小时。'));
  assert.ok(p.includes('我昨天工作八个小时。'));
  assert.ok(p.includes('"verdict"'));
  assert.ok(p.includes('close'));
});

test('parseJudgeResponse は素のJSONも前置き付きJSONも読む', () => {
  const parseJudgeResponse = extractFn('parseJudgeResponse');
  assert.deepEqual(parseJudgeResponse('{"verdict":"correct","comment":"完璧です"}'), {
    verdict: 'correct',
    comment: '完璧です',
  });
  assert.deepEqual(parseJudgeResponse('判定：{"verdict":"close","comment":"惜しい"} 以上'), {
    verdict: 'close',
    comment: '惜しい',
  });
});

test('parseJudgeResponse は不正な入力に null を返す', () => {
  const parseJudgeResponse = extractFn('parseJudgeResponse');
  assert.equal(parseJudgeResponse(''), null);
  assert.equal(parseJudgeResponse('JSONではありません'), null);
  assert.equal(parseJudgeResponse('{"verdict":"great","comment":"x"}'), null);
  assert.equal(parseJudgeResponse('{invalid json}'), null);
});

test('validateStudyInput は正常な入力を通す', () => {
  const validateStudyInput = extractFn('validateStudyInput');
  assert.deepEqual(validateStudyInput({ date: '2026-11-10', seconds: 300, source: 'auto' }), {
    date: '2026-11-10',
    seconds: 300,
    source: 'auto',
    batchId: null,
  });
  assert.deepEqual(validateStudyInput({ date: '2026-11-10', seconds: 61.6, source: 'manual' }), {
    date: '2026-11-10',
    seconds: 62,
    source: 'manual',
    batchId: null,
  });
  assert.equal(
    validateStudyInput({ date: '2026-11-10', seconds: 60, batchId: 'babc123' }).batchId,
    'babc123'
  );
  assert.equal(validateStudyInput({ date: '2026-11-10', seconds: 60, batchId: 'bad id!' }).batchId, null);
  assert.equal(validateStudyInput({ date: '2026-11-10', seconds: 60 }).source, 'auto');
});

test('validateStudyInput は不正な入力に null を返す', () => {
  const validateStudyInput = extractFn('validateStudyInput');
  assert.equal(validateStudyInput({ date: '2026/11/10', seconds: 60 }), null);
  assert.equal(validateStudyInput({ date: '', seconds: 60 }), null);
  assert.equal(validateStudyInput({ date: '2026-11-10', seconds: 0 }), null);
  assert.equal(validateStudyInput({ date: '2026-11-10', seconds: -5 }), null);
  assert.equal(validateStudyInput({ date: '2026-11-10', seconds: 36001 }), null);
  assert.equal(validateStudyInput({ date: '2026-11-10', seconds: 'abc' }), null);
  assert.equal(validateStudyInput(null), null);
});

const mTts = code.match(/function buildTtsUrl[\s\S]*?\n\}/);
assert.ok(mTts, 'buildTtsUrl not found in dist/Code.gs');
const buildTtsUrl = new Function(`${mTts[0]}; return buildTtsUrl;`)();

test('Google翻訳TTSのURLを組み立てる', () => {
  const url = buildTtsUrl('我平时搭电车上班。');
  assert.ok(url.startsWith('https://translate.google.com/translate_tts?'));
  assert.ok(url.includes('tl=zh-CN'));
  assert.ok(url.includes('client=tw-ob'));
  assert.ok(url.includes(`q=${encodeURIComponent('我平时搭电车上班。')}`));
});

test('batchSeen は送信済みバッチを見分ける', () => {
  const batchSeen = extractFn('batchSeen');
  const rows = [['b1', '2026-11-11', 60], ['b2', '2026-11-11', 120]];
  assert.equal(batchSeen(rows, 'b1'), true);
  assert.equal(batchSeen(rows, 'b3'), false);
  assert.equal(batchSeen(rows, null), false);
  assert.equal(batchSeen([], 'b1'), false);
});
