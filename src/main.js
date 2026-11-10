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

export { rowsToObjects };
