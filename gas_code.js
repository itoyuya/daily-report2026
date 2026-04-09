// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// テクニカルサポート業務日報 — Google Apps Script（閲覧用）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 使い方:
//   1. Googleスプレッドシートの「拡張機能 → Apps Script」に貼り付け
//   2. デプロイ → ウェブアプリ（実行:自分, アクセス:全員）
//   ※ PDF生成は管理用スプレッドシート（gas_code_admin.js）で行う
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ── 設定 ──────────────────────────
const CONFIG = {
  SHEET_NAME: '日報_2026',
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// doPost: フォームからのデータ受信
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
      sheet.appendRow([
        'タイムスタンプ', '日付', '氏名', 'ポスト区分',
        '開始時間', '終了時間',
        'イベント名／実施業務', '実施事項', '業務内容', '特記事項等',
        '気づき・振り返り', '勤務時間'
      ]);
    }

    // 勤務時間（時間単位）を算出
    var hours = '';
    if (data.start_time && data.end_time) {
      var st = data.start_time.split(':');
      var en = data.end_time.split(':');
      var mins = (parseInt(en[0]) * 60 + parseInt(en[1])) - (parseInt(st[0]) * 60 + parseInt(st[1]));
      if (mins < 0) mins += 24 * 60;
      hours = mins / 60;
    }

    sheet.appendRow([
      new Date(),
      data.date,
      data.member,
      data.post || '',
      data.start_time,
      data.end_time,
      data.title,
      data.tasks,
      data.content,
      data.notes,
      data.reflection,
      hours,
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
