import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// 異常系テスト：外から入ってくる入力（Sheets のセル値、学習者の回答、
// groq-gas ラッパーの応答JSON）が壊れていても、入口が例外で死なず、
// 壊れた1件が全体を巻き込まないことを固定する。
// 純粋関数は main.test.js と同じく dist/Code.gs から抜き出して評価し、
// doPost / handleTts は app.test.js と同じ GAS モック経由で検証する。
const code = readFileSync(new URL('../dist/Code.gs', import.meta.url), 'utf8');

function extractFn(name) {
  const re = new RegExp(`function ${name}[\\s\\S]*?\\n\\}`);
  const m = code.match(re);
  assert.ok(m, `${name} not found in dist/Code.gs`);
  return new Function(`${m[0]}; return ${name};`)();
}

const rowsToObjects = extractFn('rowsToObjects');
const parseJudgeResponse = extractFn('parseJudgeResponse');
const buildJudgePrompt = extractFn('buildJudgePrompt');

// JSON 経由で実際に作れる「文字列化できないオブジェクト」。
// JSON.parse('{"toString":[]}') と同じ形で、String() が TypeError を投げる。
const unstringifiable = () => ({ toString: [] });

// --- rowsToObjects：シート行の整形 ---

test('rowsToObjects は文字列の行が混ざっても他の行を返す', () => {
  const values = [
    ['id', 'zh'],
    'oops',
    ['c001', '你好'],
  ];
  assert.deepEqual(rowsToObjects(values), [{ id: 'c001', zh: '你好' }]);
});

test('rowsToObjects は数値・オブジェクトの行を読み飛ばす', () => {
  const values = [
    ['id', 'zh'],
    42,
    { id: 'c999' },
    true,
    ['c001', '你好'],
  ];
  assert.deepEqual(rowsToObjects(values), [{ id: 'c001', zh: '你好' }]);
});

test('rowsToObjects は配列でない入力に空配列を返す', () => {
  assert.deepEqual(rowsToObjects('abcd'), []);
  assert.deepEqual(rowsToObjects(42), []);
  assert.deepEqual(rowsToObjects(true), []);
});

test('rowsToObjects はヘッダー行が配列でない場合に空配列を返す', () => {
  assert.deepEqual(rowsToObjects(['id', ['c001']]), []);
  assert.deepEqual(rowsToObjects([42, ['c001']]), []);
});

test('rowsToObjects はヘッダーより長い行の余分なセルを捨てる', () => {
  assert.deepEqual(rowsToObjects([['id'], ['c001', 'extra', 'more']]), [{ id: 'c001' }]);
});

test('rowsToObjects は重複ヘッダーでは後勝ちになる（現行仕様の固定）', () => {
  assert.deepEqual(rowsToObjects([['id', 'id'], ['a', 'b']]), [{ id: 'b' }]);
});

test('rowsToObjects は 0 や false の行を空行扱いしない', () => {
  assert.deepEqual(rowsToObjects([['id', 'n'], [0, false]]), [{ id: 0, n: false }]);
});

test('rowsToObjects はオブジェクトのセル値をそのまま保持する', () => {
  const cell = unstringifiable();
  const out = rowsToObjects([['id', 'zh'], ['c001', cell]]);
  assert.equal(out.length, 1);
  assert.equal(out[0].id, 'c001');
  assert.equal(out[0].zh, cell);
});

// --- parseJudgeResponse：ラッパー応答の解釈 ---

test('parseJudgeResponse は数値・真偽値・オブジェクトの入力に null を返す', () => {
  assert.equal(parseJudgeResponse(42), null);
  assert.equal(parseJudgeResponse(true), null);
  assert.equal(parseJudgeResponse({ verdict: 'correct', comment: 'x' }), null);
});

test('parseJudgeResponse は文字列化できない comment でも判定を返す', () => {
  const text = JSON.stringify({ verdict: 'correct', comment: unstringifiable() });
  assert.deepEqual(parseJudgeResponse(text), { verdict: 'correct', comment: '' });
});

test('parseJudgeResponse はオブジェクトや配列の comment を空文字にする', () => {
  // 現行は String() で "[object Object]" や "a,b" がコメントとして漏れていた。
  assert.deepEqual(
    parseJudgeResponse(JSON.stringify({ verdict: 'close', comment: { a: 1 } })),
    { verdict: 'close', comment: '' }
  );
  assert.deepEqual(
    parseJudgeResponse(JSON.stringify({ verdict: 'close', comment: ['a', 'b'] })),
    { verdict: 'close', comment: '' }
  );
});

test('parseJudgeResponse は数値の comment を文字列化し、真偽値は空文字にする（現行仕様の固定）', () => {
  assert.deepEqual(
    parseJudgeResponse(JSON.stringify({ verdict: 'close', comment: 42 })),
    { verdict: 'close', comment: '42' }
  );
  assert.deepEqual(
    parseJudgeResponse(JSON.stringify({ verdict: 'wrong', comment: false })),
    { verdict: 'wrong', comment: '' }
  );
});

// --- buildJudgePrompt：入力は doPost が文字列化してから渡すのが前提 ---

test('buildJudgePrompt は欠けたフィールドでも文字列を返す', () => {
  assert.equal(typeof buildJudgePrompt({}), 'string');
  assert.equal(typeof buildJudgePrompt({ question: '你好吗？' }), 'string');
});

// --- app.test.js と同じ GAS モックで入口を検証する ---

const GLOBAL_NAMES = [
  'PropertiesService',
  'SpreadsheetApp',
  'CacheService',
  'ContentService',
  'UrlFetchApp',
  'Utilities',
  'ScriptApp',
];

function loadApp(mocks) {
  const factory = new Function(
    ...GLOBAL_NAMES,
    `${code}\nreturn { doGet, doPost, handleTts };`
  );
  return factory(...GLOBAL_NAMES.map((name) => mocks[name]));
}

function makeMocks() {
  const state = {
    properties: {
      SPREADSHEET_ID: 'test-spreadsheet-id',
      GROQ_API_URL: 'https://groq-wrapper.example/exec',
    },
    cache: new Map(),
    fetchCalls: [],
    fetchHandler: null,
  };
  const mocks = {
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => state.properties[key] ?? null,
      }),
    },
    SpreadsheetApp: {
      openById: () => ({
        getSheetByName: () => null,
      }),
    },
    CacheService: {
      getScriptCache: () => ({
        get: (key) => (state.cache.has(key) ? state.cache.get(key) : null),
        put: (key, value) => state.cache.set(key, String(value)),
      }),
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: (content) => ({
        content,
        mimeType: null,
        setMimeType(mime) {
          this.mimeType = mime;
          return this;
        },
      }),
    },
    UrlFetchApp: {
      fetch: (url, options) => {
        state.fetchCalls.push({ url, options });
        const res = state.fetchHandler
          ? state.fetchHandler(url, options)
          : { code: 200, text: '', content: [] };
        return {
          getResponseCode: () => res.code,
          getContentText: () => res.text ?? '',
          getContent: () => res.content ?? [],
        };
      },
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      computeDigest: (algo, text) => {
        const s = String(text);
        return Array.from({ length: 32 }, (_, i) =>
          s.length ? s.charCodeAt(i % s.length) & 0xff : 0
        );
      },
      base64Encode: (bytes) => Buffer.from(bytes).toString('base64'),
    },
    ScriptApp: {
      getProjectTriggers: () => [],
      deleteTrigger: () => {},
      newTrigger: () => {
        throw new Error('not used in these tests');
      },
    },
  };
  return { state, mocks, app: loadApp(mocks) };
}

const parseOutput = (output) => JSON.parse(output.content);

const judgeBody = (overrides = {}) =>
  JSON.stringify({
    question: '你昨天工作了几个小时？',
    qword: '几',
    expected: '我昨天工作了八个小时。',
    answer: '我昨天工作八个小时。',
    ...overrides,
  });

// --- handleTts：TTS 入口 ---

test('handleTts は文字列化できない入力を例外でなく ok:false で返す', () => {
  const { app } = makeMocks();
  assert.deepEqual(app.handleTts(unstringifiable()), { ok: false, error: 'text is empty' });
});

test('handleTts は文字列でない入力を空テキストとして扱う', () => {
  const { app, state } = makeMocks();
  assert.deepEqual(app.handleTts(42), { ok: false, error: 'text is empty' });
  assert.deepEqual(app.handleTts(null), { ok: false, error: 'text is empty' });
  assert.equal(state.fetchCalls.length, 0);
});

test('handleTts は孤立サロゲートを含むテキストを例外でなく ok:false で返す', () => {
  // buildTtsUrl の encodeURIComponent が URIError を投げるが、
  // handleTts の try/catch が受けて ok:false になる現行の振る舞いを固定する。
  const { app } = makeMocks();
  const res = app.handleTts('\uD800');
  assert.equal(res.ok, false);
  assert.equal(typeof res.error, 'string');
});

// --- doPost：判定入口とラッパー応答 ---

test('doPost はラッパーの不正な JSON 応答を api_error にする', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({ code: 200, text: 'not json at all' });
  const payload = parseOutput(app.doPost({ postData: { contents: judgeBody() } }));
  assert.deepEqual(payload, { ok: false, reason: 'api_error' });
});

test('doPost は answer が文字列でないラッパー応答を api_error にする', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({
    code: 200,
    text: JSON.stringify({ answer: 42 }),
  });
  const payload = parseOutput(app.doPost({ postData: { contents: judgeBody() } }));
  assert.deepEqual(payload, { ok: false, reason: 'api_error' });
});

test('doPost は comment が文字列化できない判定でも ok:true で返す', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({
    code: 200,
    text: JSON.stringify({
      answer: JSON.stringify({ verdict: 'correct', comment: unstringifiable() }),
    }),
  });
  const payload = parseOutput(app.doPost({ postData: { contents: judgeBody() } }));
  assert.deepEqual(payload, { ok: true, verdict: 'correct', comment: '' });
});

test('doPost は文字列化できないフィールドを bad_request にする（例外にしない）', () => {
  const { app } = makeMocks();
  const payload = parseOutput(
    app.doPost({ postData: { contents: judgeBody({ answer: unstringifiable() }) } })
  );
  assert.deepEqual(payload, { ok: false, reason: 'bad_request' });
});

test('doPost は JSON の null・配列・数値ボディを bad_request にする', () => {
  const { app } = makeMocks();
  for (const contents of ['null', '[]', '42', '"text"']) {
    const payload = parseOutput(app.doPost({ postData: { contents } }));
    assert.deepEqual(payload, { ok: false, reason: 'bad_request' });
  }
});
