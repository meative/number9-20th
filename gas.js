// Google Apps Script（スプレッドシート → 拡張機能 → Apps Script に貼る）
// デプロイ → 新しいデプロイ → 種類「ウェブアプリ」→ 実行ユーザー「自分」／アクセス「全員」→ URLを rsvp.html の ENDPOINT に貼る
// 出席者1名につき1行で記録されます（席次・名札にそのまま使えます）

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['受信日時', '出欠', '貴社名', 'ご担当者', 'お名前', 'ふりがな', '御役職', 'アレルギー']);
  }
  const d = JSON.parse(e.postData.contents);
  const now = new Date();
  const guests = (d.guests && d.guests.length) ? d.guests : [{ name: '', kana: '', title: '' }];
  guests.forEach(g => {
    sheet.appendRow([now, d.attend, d.company, d.contact, g.name, g.kana, g.title, d.allergy]);
  });
  return ContentService.createTextOutput('ok');
}
