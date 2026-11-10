# xuexi-gas

中国語レッスン復習アプリ `/xuexi` のデータAPI（Google Apps Script）。

- 正本データは Google Sheets「xuexi カードデータ」（タブ：cards / drills / mistakes）。
- スプレッドシートIDはスクリプトプロパティ `SPREADSHEET_ID` で渡す（コードに直書きしない）。
- `doGet` が3タブ分を JSON で返す。応答は CacheService で10分キャッシュする。

## 開発

- `npm run build` : `src/main.js` から `dist/Code.gs` を生成
- `npm test` : ビルド＋単体テスト

## スクリプトプロパティの設定（初回のみ）

エディタで一時関数を実行して設定する：

```js
function setProps() {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', '<シートID>');
}
```
