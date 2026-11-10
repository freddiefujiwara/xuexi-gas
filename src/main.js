// xuexi-gas: 中国語レッスン復習アプリ /xuexi のデータAPI。
// 正本は Google Sheets「xuexi カードデータ」(cards / drills / mistakes の3タブ)。
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
    mistakes: readSheetAsObjects(ss, 'mistakes'),
    generatedAt: new Date().toISOString(),
  };
}

function doGet() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('payload_v1');
  if (cached) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }
  const payload = JSON.stringify(buildPayload());
  cache.put('payload_v1', payload, 600);
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

// --- 疑問詞ドリルの自由回答を Groq で判定する ---
// APIキーはスクリプトプロパティ GROQ_API_KEY（time-table-gas と同じ流儀）。
// モデルは GROQ_MODEL で差し替え可能。キー未設定やAPI失敗時は ok:false を返し、
// フロント側は従来の自己判定にフォールバックする。

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

function judgeWithGroq(input, apiKey, model) {
  const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify({
      model: model,
      temperature: 0.2,
      max_tokens: 300,
      messages: [{ role: 'user', content: buildJudgePrompt(input) }],
    }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) return null;
  const body = JSON.parse(res.getContentText());
  const content = body && body.choices && body.choices[0] && body.choices[0].message
    ? body.choices[0].message.content
    : '';
  return parseJudgeResponse(content);
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
      const props = PropertiesService.getScriptProperties();
      const apiKey = props.getProperty('GROQ_API_KEY');
      if (!apiKey) {
        result = { ok: false, reason: 'no_key' };
      } else {
        const model = props.getProperty('GROQ_MODEL') || 'llama-3.3-70b-versatile';
        const judged = judgeWithGroq(input, apiKey, model);
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

export { rowsToObjects, buildJudgePrompt, parseJudgeResponse };
