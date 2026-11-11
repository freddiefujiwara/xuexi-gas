// xuexi-gas: 中国語レッスン復習アプリ /xuexi のデータAPI。
// 正本は Google Sheets「xuexi カードデータ」(cards / drills / words の3タブ)。
// 学習記録は端末ローカル保存に移行済みで、このAPIでは扱わない。
// スプレッドシートIDはスクリプトプロパティ SPREADSHEET_ID から読む(コードに直書きしない)。

function rowsToObjects(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row || row.every((c) => c === '' || c === null || c === undefined)) continue;
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = row[idx] === undefined ? '' : row[idx];
    });
    out.push(obj);
  }
  return out;
}

function readSheetAsObjects(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  return rowsToObjects(values);
}

function buildPayload() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_ID is not set in script properties');
  const ss = SpreadsheetApp.openById(id);
  return {
    cards: readSheetAsObjects(ss, 'cards'),
    drills: readSheetAsObjects(ss, 'drills'),
    words: readSheetAsObjects(ss, 'words'),
    generatedAt: new Date().toISOString(),
  };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.text) {
    return ContentService.createTextOutput(JSON.stringify(handleTts(e.parameter.text))).setMimeType(
      ContentService.MimeType.JSON
    );
  }
  const cache = CacheService.getScriptCache();
  const cached = cache.get('payload_v2');
  let payload;
  if (cached) {
    payload = JSON.parse(cached);
  } else {
    payload = buildPayload();
    cache.put('payload_v2', JSON.stringify(payload), 600);
  }
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// --- 発音音声（Google翻訳TTSを GAS 経由で base64 JSON 返却。/music の music-gas と同方式） ---

function buildTtsUrl(text) {
  return (
    'https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob&tl=zh-CN&q=' +
    encodeURIComponent(text)
  );
}

function ttsCacheKey(text) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text,
    Utilities.Charset.UTF_8
  );
  return (
    'tts_' +
    digest
      .map((b) => ('0' + (b & 0xff).toString(16)).slice(-2))
      .join('')
  );
}

function fetchTtsAudio(text) {
  const cache = CacheService.getScriptCache();
  const key = ttsCacheKey(text);
  const cached = cache.get(key);
  if (cached) return cached;
  const res = UrlFetchApp.fetch(buildTtsUrl(text), {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0' },
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('TTS fetch failed: ' + res.getResponseCode());
  }
  const b64 = Utilities.base64Encode(res.getContent());
  if (b64.length < 90000) cache.put(key, b64, 21600);
  return b64;
}

function handleTts(text) {
  const trimmed = String(text || '').trim();
  if (!trimmed) return { ok: false, error: 'text is empty' };
  if (trimmed.length > 200) return { ok: false, error: 'text is too long' };
  try {
    return { ok: true, mimeType: 'audio/mpeg', audio: fetchTtsAudio(trimmed) };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

// --- 疑問詞ドリルの自由回答を Groq で判定する ---
// Groq の呼び出しは既存のラッパー groq-gas（Web App）に委譲する。
// ラッパーの URL はスクリプトプロパティ GROQ_API_URL から読む。
// 未設定やAPI失敗時は ok:false を返し、フロント側は従来の自己判定にフォールバックする。

function buildJudgePrompt(input) {
  return [
    'あなたは中国語教師です。学習者が疑問詞ドリルの質問に中国語で答えました。',
    '模範解答と比べて、意味と文法の観点で判定してください。',
    '',
    `質問：${input.question}`,
    `使うべき疑問詞：${input.qword}`,
    `模範解答：${input.expected}`,
    `学習者の回答：${input.answer}`,
    '',
    '判定基準：',
    '- correct：意味が正しく伝わり、文法もほぼ正しい',
    '- close：意味は伝わるが、不自然な言い回しや小さな誤りがある',
    '- wrong：意味が違う、疑問詞への答えになっていない、文法的に大きく崩れている',
    '',
    '出力は次のJSONだけにしてください（説明文やコードフェンスは付けない）：',
    '{"verdict":"correct","comment":"日本語で一言、良かった点や直す点を書く"}',
  ].join('\n');
}

function parseJudgeResponse(text) {
  if (!text) return null;
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end <= start) return null;
  let obj = null;
  try {
    obj = JSON.parse(text.slice(start, end + 1));
  } catch (e) {
    return null;
  }
  if (!obj || ['correct', 'close', 'wrong'].indexOf(obj.verdict) < 0) return null;
  return { verdict: obj.verdict, comment: String(obj.comment || '') };
}

function judgeViaWrapper(input, wrapperUrl) {
  // The judge prompt is far beyond UrlFetchApp's URL length limit when
  // URL-encoded, so send it in the POST body (groq-gas doPost).
  const res = UrlFetchApp.fetch(wrapperUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ p: buildJudgePrompt(input) }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) return null;
  const body = JSON.parse(res.getContentText());
  if (!body || !body.answer) return null;
  return parseJudgeResponse(body.answer);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const input = {
      question: String(body.question || ''),
      qword: String(body.qword || ''),
      expected: String(body.expected || ''),
      answer: String(body.answer || ''),
    };
    if (!input.question || !input.expected || !input.answer || input.answer.length > 500) {
      result = { ok: false, reason: 'bad_request' };
    } else {
      const wrapperUrl = PropertiesService.getScriptProperties().getProperty('GROQ_API_URL');
      if (!wrapperUrl) {
        result = { ok: false, reason: 'no_wrapper' };
      } else {
        const judged = judgeViaWrapper(input, wrapperUrl);
        result = judged
          ? { ok: true, verdict: judged.verdict, comment: judged.comment }
          : { ok: false, reason: 'api_error' };
      }
    }
  } catch (err) {
    result = { ok: false, reason: 'bad_request' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
    ContentService.MimeType.JSON
  );
}

export { rowsToObjects, buildJudgePrompt, parseJudgeResponse, buildTtsUrl };
