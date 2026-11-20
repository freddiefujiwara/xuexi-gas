import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import fc from 'fast-check';

// プロパティベーステスト：生成した入力に対して常に成り立つ性質を固定する。
// 生成ドメインの考え方：
// - 文字列は書記素単位（grapheme）で生成する。実際の入力（シートのセル、
//   学習者の回答、APIのJSON文字列）は孤立サロゲートを含まないためで、
//   孤立サロゲートの振る舞いは abnormal.test.js で個別に固定している。
// - ヘッダー行はシートのセル値としてあり得るスカラー（文字列・数値・
//   真偽値・null）に限る。Sheets のセルにはオブジェクトが入らないため。
//   ただし '__proto__' は通常の代入でプロトタイプを書き換えてしまうため
//   除外する（実際のヘッダーは id / zh / ja 等の固定名）。
// - セル値そのものは JSON ドメイン全体（{"toString":[]} のような
//   文字列化できないオブジェクトを含む）まで広げる。rowsToObjects は
//   値を変換せず保持するだけなので、落とさず通せるはず、という性質を見る。
const code = readFileSync(new URL('../dist/Code.gs', import.meta.url), 'utf8');

function extractFn(name) {
  const re = new RegExp(`function ${name}[\\s\\S]*?\\n\\}`);
  const m = code.match(re);
  assert.ok(m, `${name} not found in dist/Code.gs`);
  return new Function(`${m[0]}; return ${name};`)();
}

const rowsToObjects = extractFn('rowsToObjects');
const buildJudgePrompt = extractFn('buildJudgePrompt');
const parseJudgeResponse = extractFn('parseJudgeResponse');
const buildTtsUrl = extractFn('buildTtsUrl');
const sectionCacheKey = extractFn('sectionCacheKey');
const isSectionName = extractFn('isSectionName');

const graphemeString = () => fc.string({ unit: 'grapheme' });
const headerCell = () =>
  fc
    .oneof(graphemeString(), fc.integer(), fc.boolean(), fc.constant(null))
    .filter((h) => h !== '__proto__');

test('rowsToObjects は壊れた行が混ざっても例外を投げず、オブジェクトの配列を返す', () => {
  fc.assert(
    fc.property(
      fc.array(headerCell(), { minLength: 1, maxLength: 6 }),
      fc.array(
        fc.oneof(
          fc.array(fc.jsonValue(), { maxLength: 6 }),
          fc.jsonValue(),
          graphemeString()
        ),
        { maxLength: 8 }
      ),
      (headers, rows) => {
        const out = rowsToObjects([headers, ...rows]);
        assert.ok(Array.isArray(out));
        for (const obj of out) {
          assert.equal(typeof obj, 'object');
          assert.notEqual(obj, null);
          assert.ok(!Array.isArray(obj));
        }
      }
    ),
    { numRuns: 100 }
  );
});

test('rowsToObjects は行の保存則：空でない配列行が過不足なくオブジェクトになる', () => {
  fc.assert(
    fc.property(
      fc.uniqueArray(graphemeString().filter((s) => s !== '' && s !== '__proto__'), {
        minLength: 1,
        maxLength: 5,
      }),
      fc.array(fc.array(fc.jsonValue(), { maxLength: 5 }), { maxLength: 8 }),
      (headers, rows) => {
        const out = rowsToObjects([headers, ...rows]);
        const kept = rows.filter(
          (row) => !row.every((c) => c === '' || c === null || c === undefined)
        );
        assert.equal(out.length, kept.length);
        kept.forEach((row, i) => {
          assert.deepEqual(Object.keys(out[i]).sort(), [...headers].sort());
          headers.forEach((h, idx) => {
            assert.deepEqual(out[i][h], row[idx] === undefined ? '' : row[idx]);
          });
        });
      }
    ),
    { numRuns: 100 }
  );
});

test('parseJudgeResponse はどんな入力でも例外を投げず、null か {verdict, comment文字列} を返す', () => {
  fc.assert(
    fc.property(
      fc.oneof(graphemeString(), fc.jsonValue(), fc.constant(null), fc.constant(undefined)),
      (input) => {
        const out = parseJudgeResponse(input);
        if (out === null) return;
        assert.ok(['correct', 'close', 'wrong'].includes(out.verdict));
        assert.equal(typeof out.comment, 'string');
      }
    ),
    { numRuns: 100 }
  );
});

test('parseJudgeResponse は verdict と comment の往復が保たれる', () => {
  fc.assert(
    fc.property(
      fc.constantFrom('correct', 'close', 'wrong'),
      graphemeString(),
      (verdict, comment) => {
        const text = JSON.stringify({ verdict, comment });
        assert.deepEqual(parseJudgeResponse(text), { verdict, comment });
      }
    ),
    { numRuns: 100 }
  );
});

test('buildJudgePrompt は文字列フィールドをすべてそのまま含む', () => {
  fc.assert(
    fc.property(
      graphemeString(),
      graphemeString(),
      graphemeString(),
      graphemeString(),
      (question, qword, expected, answer) => {
        const p = buildJudgePrompt({ question, qword, expected, answer });
        assert.equal(typeof p, 'string');
        for (const field of [question, qword, expected, answer]) {
          assert.ok(p.includes(field));
        }
      }
    ),
    { numRuns: 100 }
  );
});

test('buildTtsUrl は整った文字列では常に encodeURIComponent した URL になる', () => {
  fc.assert(
    fc.property(graphemeString(), (text) => {
      const url = buildTtsUrl(text);
      assert.ok(url.startsWith('https://translate.google.com/translate_tts?'));
      assert.ok(url.endsWith(`q=${encodeURIComponent(text)}`));
    }),
    { numRuns: 100 }
  );
});

test('isSectionName は3つのセクション名だけを通し、キャッシュキーは接頭辞付きで安定する', () => {
  fc.assert(
    fc.property(graphemeString(), (name) => {
      assert.equal(isSectionName(name), ['cards', 'drills', 'words'].includes(name));
      assert.equal(sectionCacheKey(name), `section_v1_${name}`);
    }),
    { numRuns: 100 }
  );
});

// --- doPost の全域性：どんな JSON ボディでも例外なく {ok:真偽値} の JSON を返す ---

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
    `${code}\nreturn { doPost };`
  );
  return factory(...GLOBAL_NAMES.map((name) => mocks[name]));
}

function makeJudgeMocks(fetchText) {
  const state = { cache: new Map() };
  const mocks = {
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) =>
          ({
            SPREADSHEET_ID: 'test-spreadsheet-id',
            GROQ_API_URL: 'https://groq-wrapper.example/exec',
          })[key] ?? null,
      }),
    },
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => null }) },
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
        setMimeType() {
          return this;
        },
      }),
    },
    UrlFetchApp: {
      fetch: () => ({
        getResponseCode: () => 200,
        getContentText: () => fetchText,
        getContent: () => [],
      }),
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      computeDigest: () => new Array(32).fill(0),
      base64Encode: (bytes) => Buffer.from(bytes).toString('base64'),
    },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger: () => {}, newTrigger: () => {} },
  };
  return loadApp(mocks);
}

test('doPost は生成した JSON ボディとラッパー応答でも例外を投げず ok 真偽値を返す', () => {
  fc.assert(
    fc.property(fc.jsonValue(), graphemeString(), (body, wrapperText) => {
      const app = makeJudgeMocks(wrapperText);
      const output = app.doPost({
        postData: { contents: JSON.stringify(body) ?? 'null' },
      });
      const payload = JSON.parse(output.content);
      assert.equal(typeof payload.ok, 'boolean');
      if (payload.ok) {
        assert.ok(['correct', 'close', 'wrong'].includes(payload.verdict));
        assert.equal(typeof payload.comment, 'string');
      } else {
        assert.ok(['bad_request', 'no_wrapper', 'api_error'].includes(payload.reason));
      }
    }),
    { numRuns: 100 }
  );
});
