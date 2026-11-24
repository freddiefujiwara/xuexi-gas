# xuexi-gas

中国語レッスン復習アプリ `/xuexi` のデータAPI（Google Apps Script）。

- 正本データは Google Sheets「xuexi」（タブ：cards / drills / words）。
- スプレッドシートIDはスクリプトプロパティ `SPREADSHEET_ID` で渡す（コードに直書きしない）。
- `doGet` が3タブ分を JSON で返す。応答は CacheService で6時間キャッシュし、1時間ごとのトリガー（`preCacheXuexi`）が温め直す。

## 開発

- `npm run build` : `src/main.js` から `dist/Code.gs` を生成
- `npm test` : ビルド＋単体テスト

## スクリプトプロパティの設定（初回のみ）

エディタで一時関数を実行して設定する：

```js
function setProps() {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', '<シートID>');
  // 判定（Groq ラッパー経由）に使う。未設定のときは判定がフェイルクローズする
  PropertiesService.getScriptProperties().setProperty('GROQ_API_URL', '<groq-gas の /exec URL>');
  PropertiesService.getScriptProperties().setProperty('GROQ_GAS_TOKEN', '<groq-gas の API_TOKEN>');
}
```
