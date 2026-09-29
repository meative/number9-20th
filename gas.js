// Google Apps Script（スプレッドシート → 拡張機能 → Apps Script に貼る）
// デプロイ → 新しいデプロイ → 種類「ウェブアプリ」→ 実行ユーザー「自分」／アクセス「全員」→ URLを index.html と staff/index.html の ENDPOINT に貼る
// 取引先フォーム → シート「取引先」（出席者1名につき1行）
// スタッフフォーム → シート「スタッフ」（1名1行、お子様の希望人数・年齢つき）

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const d = JSON.parse(e.postData.contents);
  const now = new Date();

  if (d.form === 'staff') {
    const sh = getSheet(ss, 'スタッフ', ['受信日時', '出欠', '所属', 'お名前', 'ふりがな', 'お子様参加希望', 'お子様人数', 'お子様年齢', 'アレルギー', '抽選結果']);
    sh.appendRow([now, d.attend, d.group, d.name, d.kana, d.kids, d.kidsCount, d.kidsAges, d.allergy, '']);
  } else {
    const sh = getSheet(ss, '取引先', ['受信日時', '出欠', '貴社名', 'ご担当者', 'お名前', 'ふりがな', '御役職', 'アレルギー']);
    const guests = (d.guests && d.guests.length) ? d.guests : [{ name: '', kana: '', title: '' }];
    guests.forEach(g => sh.appendRow([now, d.attend, d.company, d.contact, g.name, g.kana, g.title, d.allergy]));
  }
  return ContentService.createTextOutput('ok');
}

function getSheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  if (sh.getLastRow() === 0) { sh.appendRow(header); }
  return sh;
}
