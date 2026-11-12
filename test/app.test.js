import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// dist/Code.gs 全体を、GAS グローバルのモックを引数として読み込む。
// 純粋関数だけを抜き出す main.test.js と違い、こちらは関数同士の呼び出し
// （doGet → readSection → readSheetAsObjects など）もそのまま動かせるので、
// 本番入口と GAS 依存層の検証に使う。
const code = readFileSync(new URL('../dist/Code.gs', import.meta.url), 'utf8');

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
    `${code}\nreturn { doGet, doPost, readSection, readSheetAsObjects, buildPayload, sectionsBuiltAt, preCacheXuexi, setupCacheWarmer, handleTts, fetchTtsAudio, judgeViaWrapper };`
  );
  return factory(...GLOBAL_NAMES.map((name) => mocks[name]));
}

function makeMocks() {
  const state = {
    properties: {
      SPREADSHEET_ID: 'test-spreadsheet-id',
      GROQ_API_URL: 'https://groq-wrapper.example/exec',
    },
    sheets: {
      cards: [
        ['id', 'zh', 'ja'],
        ['c001', '你好', 'こんにちは'],
      ],
      drills: [
        ['id', 'q'],
        ['d001', '你昨天工作了几个小时？'],
      ],
      words: [
        ['id', 'zh'],
        ['w001', '谢谢'],
      ],
    },
    cache: new Map(),
    sheetReads: {},
    fetchCalls: [],
    fetchHandler: null,
    triggers: [],
    deletedTriggers: [],
    createdTriggers: [],
  };

  const mocks = {
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => state.properties[key] ?? null,
      }),
    },
    SpreadsheetApp: {
      openById: () => ({
        getSheetByName: (name) => {
          if (!state.sheets[name]) return null;
          return {
            getDataRange: () => ({
              getValues: () => {
                state.sheetReads[name] = (state.sheetReads[name] || 0) + 1;
                return state.sheets[name];
              },
            }),
          };
        },
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
      getProjectTriggers: () => state.triggers,
      deleteTrigger: (trigger) => state.deletedTriggers.push(trigger),
      newTrigger: (handler) => {
        const record = { handler, hours: null };
        return {
          timeBased() {
            return this;
          },
          everyHours(hours) {
            record.hours = hours;
            return this;
          },
          create() {
            state.createdTriggers.push(record);
            return record;
          },
        };
      },
    },
  };

  return { state, mocks, app: loadApp(mocks) };
}

const parseOutput = (output) => JSON.parse(output.content);

// --- シート読み取りとセクションキャッシュ ---

test('readSheetAsObjects はシートを行オブジェクトに変換し、無いシートは空配列', () => {
  const { app, mocks } = makeMocks();
  const ss = mocks.SpreadsheetApp.openById('any');
  assert.deepEqual(app.readSheetAsObjects(ss, 'cards'), [
    { id: 'c001', zh: '你好', ja: 'こんにちは' },
  ]);
  assert.deepEqual(app.readSheetAsObjects(ss, 'unknown'), []);
});

test('readSection はシートを読んでキャッシュし、2回目は再読しない', () => {
  const { app, state } = makeMocks();
  const first = app.readSection('cards', false);
  const second = app.readSection('cards', false);
  assert.deepEqual(first, [{ id: 'c001', zh: '你好', ja: 'こんにちは' }]);
  assert.deepEqual(second, first);
  assert.equal(state.sheetReads.cards, 1);
  assert.ok(state.cache.has('section_v1_cards'));
  assert.ok(state.cache.has('sections_built_at'));
});

test('readSection は forceRefresh でキャッシュを無視して読み直す', () => {
  const { app, state } = makeMocks();
  app.readSection('cards', false);
  state.sheets.cards = [
    ['id', 'zh', 'ja'],
    ['c002', '再见', 'さようなら'],
  ];
  const refreshed = app.readSection('cards', true);
  assert.deepEqual(refreshed, [{ id: 'c002', zh: '再见', ja: 'さようなら' }]);
  assert.equal(state.sheetReads.cards, 2);
});

test('sectionsBuiltAt は保存済みの生成時刻を返し、無ければ現在時刻を返す', () => {
  const { app, state } = makeMocks();
  const fallback = app.sectionsBuiltAt();
  assert.ok(Math.abs(Date.parse(fallback) - Date.now()) < 5000);
  state.cache.set('sections_built_at', '2026-01-01T00:00:00.000Z');
  assert.equal(app.sectionsBuiltAt(), '2026-01-01T00:00:00.000Z');
});

test('buildPayload は3セクションと generatedAt を返す', () => {
  const { app } = makeMocks();
  const payload = app.buildPayload(false);
  assert.deepEqual(payload.cards, [{ id: 'c001', zh: '你好', ja: 'こんにちは' }]);
  assert.deepEqual(payload.drills, [{ id: 'd001', q: '你昨天工作了几个小时？' }]);
  assert.deepEqual(payload.words, [{ id: 'w001', zh: '谢谢' }]);
  assert.equal(typeof payload.generatedAt, 'string');
});

test('SPREADSHEET_ID 未設定なら readSection はエラーを投げる', () => {
  const { app, state } = makeMocks();
  delete state.properties.SPREADSHEET_ID;
  assert.throws(() => app.readSection('cards', false), /SPREADSHEET_ID is not set/);
});

test('preCacheXuexi は全セクションを強制再読してキャッシュを更新する', () => {
  const { app, state } = makeMocks();
  app.buildPayload(false);
  state.sheets.words = [
    ['id', 'zh'],
    ['w002', '请'],
  ];
  app.preCacheXuexi();
  assert.deepEqual(app.readSection('words', false), [{ id: 'w002', zh: '请' }]);
  assert.equal(state.sheetReads.cards, 2);
  assert.equal(state.sheetReads.drills, 2);
  assert.equal(state.sheetReads.words, 2);
});

// --- doGet ---

test('doGet はパラメータ無しで全セクションを JSON で返す', () => {
  const { app } = makeMocks();
  const output = app.doGet({ parameter: {} });
  assert.equal(output.mimeType, 'application/json');
  const payload = parseOutput(output);
  assert.deepEqual(payload.cards, [{ id: 'c001', zh: '你好', ja: 'こんにちは' }]);
  assert.deepEqual(payload.drills, [{ id: 'd001', q: '你昨天工作了几个小时？' }]);
  assert.deepEqual(payload.words, [{ id: 'w001', zh: '谢谢' }]);
  assert.equal(typeof payload.generatedAt, 'string');
});

test('doGet は type 指定でそのセクションだけを返す', () => {
  const { app } = makeMocks();
  const payload = parseOutput(app.doGet({ parameter: { type: 'cards' } }));
  assert.deepEqual(payload.cards, [{ id: 'c001', zh: '你好', ja: 'こんにちは' }]);
  assert.equal(payload.drills, undefined);
  assert.equal(payload.words, undefined);
  assert.equal(typeof payload.generatedAt, 'string');
});

test('doGet は未知の type では全ペイロードにフォールバックする', () => {
  const { app } = makeMocks();
  const payload = parseOutput(app.doGet({ parameter: { type: 'bogus' } }));
  assert.ok(payload.cards);
  assert.ok(payload.drills);
  assert.ok(payload.words);
});

// --- TTS ---

test('doGet は text パラメータを TTS として処理する', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({ code: 200, content: [1, 2, 3] });
  const payload = parseOutput(app.doGet({ parameter: { text: ' 你好 ' } }));
  assert.deepEqual(payload, {
    ok: true,
    mimeType: 'audio/mpeg',
    audio: Buffer.from([1, 2, 3]).toString('base64'),
  });
  assert.equal(state.fetchCalls.length, 1);
  assert.ok(state.fetchCalls[0].url.startsWith('https://translate.google.com/translate_tts?'));
  assert.ok(state.fetchCalls[0].url.includes(`q=${encodeURIComponent('你好')}`));
});

test('handleTts は空白のみのテキストと長すぎるテキストを拒否する', () => {
  const { app, state } = makeMocks();
  assert.deepEqual(app.handleTts('   '), { ok: false, error: 'text is empty' });
  assert.deepEqual(app.handleTts('あ'.repeat(201)), {
    ok: false,
    error: 'text is too long',
  });
  assert.equal(state.fetchCalls.length, 0);
});

test('fetchTtsAudio は同じ文をキャッシュから返し、取得は1回だけ', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({ code: 200, content: [9, 8, 7] });
  const first = app.fetchTtsAudio('谢谢');
  const second = app.fetchTtsAudio('谢谢');
  assert.equal(first, Buffer.from([9, 8, 7]).toString('base64'));
  assert.equal(second, first);
  assert.equal(state.fetchCalls.length, 1);
});

test('fetchTtsAudio は大きな音声をキャッシュしない', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({ code: 200, content: new Array(70000).fill(7) });
  app.fetchTtsAudio('你好');
  app.fetchTtsAudio('你好');
  assert.equal(state.fetchCalls.length, 2);
});

test('handleTts は TTS 取得の失敗を ok:false で返す', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({ code: 503, text: 'down' });
  assert.deepEqual(app.handleTts('你好'), {
    ok: false,
    error: 'TTS fetch failed: 503',
  });
});

// --- doPost（Groq 判定） ---

const judgeBody = (overrides = {}) =>
  JSON.stringify({
    question: '你昨天工作了几个小时？',
    qword: '几',
    expected: '我昨天工作了八个小时。',
    answer: '我昨天工作八个小时。',
    ...overrides,
  });

test('doPost は不正な JSON ボディを bad_request にする', () => {
  const { app } = makeMocks();
  const payload = parseOutput(app.doPost({ postData: { contents: '{not json' } }));
  assert.deepEqual(payload, { ok: false, reason: 'bad_request' });
});

test('doPost は必須項目の欠落と長すぎる回答を bad_request にする', () => {
  const { app } = makeMocks();
  assert.deepEqual(
    parseOutput(app.doPost({ postData: { contents: judgeBody({ answer: '' }) } })),
    { ok: false, reason: 'bad_request' }
  );
  assert.deepEqual(
    parseOutput(app.doPost({ postData: { contents: judgeBody({ answer: 'a'.repeat(501) }) } })),
    { ok: false, reason: 'bad_request' }
  );
});

test('doPost は GROQ_API_URL 未設定なら no_wrapper を返す', () => {
  const { app, state } = makeMocks();
  delete state.properties.GROQ_API_URL;
  const payload = parseOutput(app.doPost({ postData: { contents: judgeBody() } }));
  assert.deepEqual(payload, { ok: false, reason: 'no_wrapper' });
  assert.equal(state.fetchCalls.length, 0);
});

test('doPost はラッパーの判定結果を ok:true で返す', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({
    code: 200,
    text: JSON.stringify({
      answer: '判定：{"verdict":"close","comment":"惜しい"}',
    }),
  });
  const payload = parseOutput(app.doPost({ postData: { contents: judgeBody() } }));
  assert.deepEqual(payload, { ok: true, verdict: 'close', comment: '惜しい' });
  assert.equal(state.fetchCalls.length, 1);
  assert.equal(state.fetchCalls[0].url, 'https://groq-wrapper.example/exec');
  assert.equal(state.fetchCalls[0].options.method, 'post');
  const sent = JSON.parse(state.fetchCalls[0].options.payload);
  assert.ok(sent.p.includes('你昨天工作了几个小时？'));
  assert.ok(sent.p.includes('我昨天工作八个小时。'));
});

test('doPost はラッパーの HTTP 失敗を api_error にする', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({ code: 500, text: 'error' });
  const payload = parseOutput(app.doPost({ postData: { contents: judgeBody() } }));
  assert.deepEqual(payload, { ok: false, reason: 'api_error' });
});

test('doPost は answer の無いラッパー応答を api_error にする', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({ code: 200, text: JSON.stringify({ cached: true }) });
  const payload = parseOutput(app.doPost({ postData: { contents: judgeBody() } }));
  assert.deepEqual(payload, { ok: false, reason: 'api_error' });
});

test('doPost は判定 JSON を含まない answer を api_error にする', () => {
  const { app, state } = makeMocks();
  state.fetchHandler = () => ({
    code: 200,
    text: JSON.stringify({ answer: '判定できませんでした' }),
  });
  const payload = parseOutput(app.doPost({ postData: { contents: judgeBody() } }));
  assert.deepEqual(payload, { ok: false, reason: 'api_error' });
});

// --- キャッシュ温め直しトリガー ---

test('setupCacheWarmer は既存の preCacheXuexi トリガーだけを消して1時間ごとに作り直す', () => {
  const { app, state } = makeMocks();
  const warmer = { getHandlerFunction: () => 'preCacheXuexi' };
  const other = { getHandlerFunction: () => 'someOtherJob' };
  state.triggers = [warmer, other];
  app.setupCacheWarmer();
  assert.deepEqual(state.deletedTriggers, [warmer]);
  assert.deepEqual(state.createdTriggers, [{ handler: 'preCacheXuexi', hours: 1 }]);
});
