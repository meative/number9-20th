// Google Apps Script
// 1) Googleスプレッドシートを新規作成 → 拡張機能 → Apps Script → このコードを貼って保存
// 2) デプロイ → 新しいデプロイ → 種類「ウェブアプリ」→ 実行ユーザー「自分」／アクセス「全員」→ デプロイ
// 3) 出てきたURL（…/exec）を index.html と staff/index.html の ENDPOINT = "" に貼る
// 4) 一度「setup」関数を実行すると、空のシート3枚（取引先／スタッフ／集計）と集計式ができます
//
// 取引先フォーム → シート「取引先」（出席者1名につき1行）
// スタッフフォーム → シート「スタッフ」（1名1行、お子様の希望人数・年齢つき）
// 集計 → シート「集計」（自動計算）
// エンドロール → 各シートの「エンドロール表記名」「エンドロール一言」列に記入 → 映像生成に使用

const H_CO = ['受信日時', '出欠', '貴社名', 'お名前', 'ふりがな', '御役職', 'アレルギー', '卓番', 'エンドロール表記名', 'エンドロール一言'];
const H_ST = ['受信日時', '出欠', '所属', 'お名前', 'ふりがな', 'お子様参加希望', 'お子様人数', 'お子様年齢', 'アレルギー', '抽選結果', '卓番', 'エンドロール表記名', 'エンドロール一言'];

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const d = JSON.parse(e.postData.contents);
  const now = new Date();
  if (d.form === 'staff') {
    const sh = getSheet(ss, 'スタッフ', H_ST);
    sh.appendRow([now, d.attend, d.group, d.name, d.kana, d.kids, d.kidsCount, d.kidsAges, d.allergy, '', '', d.name, '']);
  } else {
    const sh = getSheet(ss, '取引先', H_CO);
    const guests = (d.guests && d.guests.length) ? d.guests : [{ name: '', kana: '', title: '' }];
    guests.forEach(g => sh.appendRow([now, d.attend, d.company, g.name, g.kana, g.title, d.allergy, '', g.name, '']));
  }
  ensureSummary(ss);
  return ContentService.createTextOutput('ok');
}

// ▼ 全データを消して作り直す（テスト後のリセット用）。本番開始後は実行しないこと
function reset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tmp = ss.insertSheet('_tmp');
  ['取引先', 'スタッフ', '集計'].forEach(n => { const sh = ss.getSheetByName(n); if (sh) ss.deleteSheet(sh); });
  setup();
  ss.deleteSheet(tmp);
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getSheet(ss, '取引先', H_CO);
  getSheet(ss, 'スタッフ', H_ST);
  ensureSummary(ss);
  const def = ss.getSheetByName('シート1'); if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  ss.setActiveSheet(ss.getSheetByName('集計')); ss.moveActiveSheet(1);
}

function getSheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#1a1918').setFontColor('#f4f2ed');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
  }
  return sh;
}

function ensureSummary(ss) {
  if (ss.getSheetByName('集計')) return;
  const sh = ss.insertSheet('集計');
  const rows = [
    ['■ 全体', '', ''],
    ['区分', '出席（人）', '欠席（件）'],
    ['取引先', '=COUNTIF(取引先!B:B,"出席")', '=COUNTIF(取引先!B:B,"欠席")'],
    ['スタッフ', '=COUNTIF(スタッフ!B:B,"出席")', '=COUNTIF(スタッフ!B:B,"欠席")'],
    ['お子様（希望・抽選前）', '=SUMPRODUCT((スタッフ!F:F="希望する")*(スタッフ!B:B="出席")*IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(スタッフ!G:G,"名以上",""),"名","")),0))', ''],
    ['お子様（当選）', '=COUNTIF(スタッフ!J:J,"当選")', ''],
    ['合計（大人）', '=B3+B4', ''],
    ['', '', ''],
    ['■ 取引先：会社別（出席者数）', '', ''],
    ['会社名', '出席（人）', ''],
    ['=IFERROR(SORT(UNIQUE(FILTER(取引先!C:C,取引先!B:B="出席",取引先!C:C<>""))),"")', '=ARRAYFORMULA(IF(A11:A60="","",COUNTIFS(取引先!C:C,A11:A60,取引先!B:B,"出席")))', ''],
  ];
  for (let i = 0; i < 49; i++) rows.push(['', '', '']);
  rows.push(['', '', '']);
  rows.push(['■ スタッフ：所属別（出席者数）', '', '']);
  rows.push(['所属', '出席（人）', 'お子様希望（人）']);
  const groups = ['美容（WOODSTOCK）','美容（JANIS）','保育（STRAWBERRY FIELD 月寒）','保育（STRAWBERRY FIELD 白石）','保育（STRAWBERRY FIELD サッポロファクトリー）','保育（ストロベリーフィールド幼稚舎）','飲食・食品（Meative）','福祉（元町ベース）','本部（Number9 / JOINTS NINE）'];
  groups.forEach(g => rows.push([g, `=COUNTIFS(スタッフ!C:C,"${g}",スタッフ!B:B,"出席")`, `=COUNTIFS(スタッフ!C:C,"${g}",スタッフ!B:B,"出席",スタッフ!F:F,"希望する")`]));
  rows.push(['', '', '']);
  rows.push(['■ アレルギー申告（出席者のみ）', '', '']);
  rows.push(['=IFERROR(FILTER({取引先!C:C,取引先!D:D,取引先!G:G},取引先!B:B="出席",取引先!G:G<>""),"なし")', '', '']);
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.setColumnWidth(1, 300); sh.setColumnWidths(2, 2, 120);
  [1, 9, 62, 62 + 2 + groups.length + 1].forEach(r => sh.getRange(r, 1).setFontWeight('bold'));
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);
}
