// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// テクニカルサポート業務日報 — Google Apps Script
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 使い方:
//   1. Googleスプレッドシートの「拡張機能 → Apps Script」に貼り付け
//   2. CONFIG の DRIVE_FOLDER_ID と COMPANY_NAME を設定
//   3. デプロイ → ウェブアプリ（実行:自分, アクセス:全員）
//   4. PDF生成: generateDailyReport("2026-04-06") を実行
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ── 設定 ──────────────────────────
const CONFIG = {
  SHEET_NAME: '日報_2026',
  DRIVE_FOLDER_ID: 'YOUR_FOLDER_ID',   // ← Google DriveフォルダIDに変更
  COMPANY_NAME: 'arsaffix Inc.',
  RESPONSIBLE_PERSON: '伊藤友哉（arsaffix Inc.）',
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
        'タイムスタンプ', '日付', '氏名',
        '開始時間', '終了時間',
        'イベント名／実施業務', '実施事項', '業務内容', '特記事項等',
        '気づき・振り返り', '勤務時間'
      ]);
    }

    // 勤務時間（時間単位）を算出
    var hours = '';
    if (data.start_time && data.end_time) {
      var s = data.start_time.split(':');
      var e = data.end_time.split(':');
      var mins = (parseInt(e[0]) * 60 + parseInt(e[1])) - (parseInt(s[0]) * 60 + parseInt(s[1]));
      if (mins < 0) mins += 24 * 60;
      hours = mins / 60;
    }

    sheet.appendRow([
      new Date(),
      data.date,
      data.member,
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

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// generateDailyReport: 指定日のPDFを生成しDriveに保存
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 使い方: generateDailyReport("2026-04-06")
function generateDailyReport(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('「日報_2026」シートが見つかりません');

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(row => row[1] === dateStr);

  if (rows.length === 0) {
    throw new Error(dateStr + ' のデータが見つかりません');
  }

  // 日付を和暦に変換
  const date = new Date(dateStr + 'T00:00:00');
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  const reiwa = date.getFullYear() - 2018;
  const dateDisplay = '令和' + reiwa + '年' + (date.getMonth() + 1) + '月' + date.getDate() + '日（' + weekdays[date.getDay()] + '）';

  // 全メンバー分を1ページにcombine
  // row: [タイムスタンプ, 日付, 氏名, 開始, 終了, タイトル, 実施事項, 内容, 特記事項等, 気づき・振り返り, 勤務時間]
  //        0            1     2     3     4     5        6        7      8            9             10
  var shifts = rows.map(function(r) { return (r[2] || '') + '：' + (r[3] || '') + '〜' + (r[4] || '') + '（' + r[10] + 'h）'; }).join('\n');
  var titles = rows.map(function(r) { return r[5] || ''; }).filter(Boolean).join('\n');
  var combineField = function(idx) {
    return rows.map(function(r) {
      if (!r[idx]) return null;
      var parts = [];
      if (rows.length > 1) parts.push('【' + (r[2] || '') + '】');
      parts.push(r[idx]);
      return parts.join('\n');
    }).filter(Boolean).join('\n\n');
  };
  var tasks = combineField(6);
  var contents = combineField(7);
  var notes = combineField(8);

  var page = buildPage({
    date: dateDisplay,
    shift: shifts,
    title: titles,
    tasks: tasks,
    content: contents,
    notes: notes,
    responsible: CONFIG.RESPONSIBLE_PERSON,
    company: CONFIG.COMPANY_NAME,
  });

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<style>' + getReportCss() + '</style>'
    + '</head><body>' + page + '</body></html>';

  // PDF生成
  var blob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .setName('業務日報_' + dateStr + '.pdf');

  // Driveに保存（同名ファイルは上書き）
  var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  var existing = folder.getFilesByName('業務日報_' + dateStr + '.pdf');
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  var file = folder.createFile(blob);
  Logger.log('PDF保存完了: ' + file.getUrl());
  return file.getUrl();
}

// ── 1ページ分のHTML ──────────────────────
function buildPage(d) {
  var nl2br = function(str) {
    return (str || '').replace(/\n/g, '<br>');
  };

  return ''
    + '<div class="page">'
    + '  <div class="doc-id">別紙６</div>'
    + '  <h1>テクニカルサポート業務日報</h1>'
    + ''
    + '  <div class="header-row">'
    + '    <div class="date">' + d.date + '</div>'
    + '    <table class="stamp-table">'
    + '      <tr><td colspan="4" class="stamp-label">※押印欄</td></tr>'
    + '      <tr><td class="stamp-cell"></td><td class="stamp-cell"></td><td class="stamp-cell"></td><td class="stamp-cell"></td></tr>'
    + '    </table>'
    + '  </div>'
    + ''
    + '  <table class="report-table">'
    + '    <tr>'
    + '      <th>イベント名／<br>実施業務</th>'
    + '      <td>' + nl2br(d.title) + '</td>'
    + '    </tr>'
    + '    <tr>'
    + '      <th>実施事項</th>'
    + '      <td>' + nl2br(d.tasks) + '</td>'
    + '    </tr>'
    + '    <tr>'
    + '      <th>業務内容</th>'
    + '      <td class="content-cell">' + nl2br(d.content) + '</td>'
    + '    </tr>'
    + '    <tr>'
    + '      <th>勤務シフト</th>'
    + '      <td>' + nl2br(d.shift) + '</td>'
    + '    </tr>'
    + '    <tr>'
    + '      <th>特記事項等</th>'
    + '      <td>' + nl2br(d.notes) + '</td>'
    + '    </tr>'
    + '    <tr>'
    + '      <th>責任者氏名</th>'
    + '      <td>' + d.responsible + '</td>'
    + '    </tr>'
    + '  </table>'
    + ''
    + '  <div class="company">' + d.company + '</div>'
    + '</div>';
}

// ── PDF用CSS ─────────────────────────
function getReportCss() {
  return ''
    + '@page { size: A4 portrait; margin: 20mm 15mm 20mm 15mm; }'
    + 'body { font-family: "Noto Sans JP", "Hiragino Kaku Gothic Pro", "Yu Gothic", sans-serif; font-size: 11pt; color: #000; margin: 0; padding: 0; }'
    + '.page { position: relative; }'
    + ''
    + '.doc-id { text-align: right; font-size: 10pt; margin-bottom: 8px; }'
    + 'h1 { text-align: center; font-size: 16pt; font-weight: bold; margin-bottom: 20px; }'
    + ''
    + '.header-row { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; }'
    + '.date { font-size: 11pt; padding-top: 4px; }'
    + ''
    + '.stamp-table { border-collapse: collapse; }'
    + '.stamp-label { font-size: 9pt; text-align: left; padding: 2px 4px; border: none; }'
    + '.stamp-cell { width: 28mm; height: 20mm; border: 1px solid #000; }'
    + ''
    + '.report-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }'
    + '.report-table th, .report-table td { border: 1px solid #000; padding: 10px 12px; vertical-align: top; font-size: 10.5pt; line-height: 1.7; }'
    + '.report-table th { width: 22%; font-weight: bold; background: #f8f8f8; text-align: left; white-space: nowrap; }'
    + '.report-table td { width: 78%; }'
    + '.report-table .content-cell { min-height: 180px; }'
    + ''
    + '.company { text-align: center; font-size: 11pt; margin-top: 16px; }';
}
