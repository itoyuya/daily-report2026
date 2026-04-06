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
  DRIVE_FOLDER_ID: '1nqFYO3mjOcwupn--p1Q7Bh8y-SAtTOl5',   // ← Google DriveフォルダIDに変更
  COMPANY_NAME: 'arsaffix Inc.',
  RESPONSIBLE_PERSON: '伊藤友哉（arsaffix Inc.）',
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// メニュー: スプレッドシートを開いたときにカスタムメニューを追加
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報PDF')
    .addItem('PDF出力（A1セルの年月を使用）', 'runFromSheet')
    .addToUi();
}

function runFromSheet() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');

  if (!sheet) {
    // 「設定」シートがなければ作成
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('設定');
    sheet.getRange('A1').setValue('出力対象');
    sheet.getRange('B1').setValue('2026-04');
    sheet.getRange('A3').setValue('↑ B1セルにYYYY-MM または YYYY-MM-DD を入力し、メニュー「日報PDF → PDF出力」を実行');
    ui.alert('「設定」シートを作成しました。\nB1セルに年月（例: 2026-04）を入力してから再度実行してください。');
    return;
  }

  var dateStr = String(sheet.getRange('B1').getValue()).trim();

  if (!dateStr || (!/^\d{4}-\d{2}$/.test(dateStr) && !/^\d{4}-\d{2}-\d{2}$/.test(dateStr))) {
    ui.alert('「設定」シートのB1セルに YYYY-MM または YYYY-MM-DD の形式で入力してください。\n例: 2026-04');
    return;
  }

  ui.alert('PDF出力を開始します: ' + dateStr);

  try {
    var result = generateDailyReport(dateStr);
    if (Array.isArray(result)) {
      ui.alert('完了: ' + result.length + '日分のPDFを出力しました。');
    } else {
      ui.alert('完了: PDFを出力しました。\n' + result);
    }
  } catch (err) {
    ui.alert('エラー: ' + err.message);
  }
}

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

// ── 日付をYYYY-MM-DD文字列に変換 ──────────────────
function toDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = String(val.getMonth() + 1).padStart(2, '0');
    var d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(val);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// generateDailyReport: 指定日 or 指定月のPDFを生成しDriveに保存
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 使い方:
//   generateDailyReport("2026-04-06")  ← 1日分
//   generateDailyReport("2026-04")     ← その月の全日分を日ごとに生成
function generateDailyReport(dateStr) {
  // YYYY-MM 形式の場合は月一括処理
  if (/^\d{4}-\d{2}$/.test(dateStr)) {
    return generateMonthlyReports(dateStr);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('「日報_2026」シートが見つかりません');

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(row => toDateStr(row[1]) === dateStr);

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
  // 勤務シフト: 時間のみ表示（HH:MM形式）
  var formatTime = function(val) {
    if (!val) return '';
    var s = String(val);
    // Date型の場合はHH:MMを抽出
    if (val instanceof Date) {
      return String(val.getHours()).padStart(2, '0') + ':' + String(val.getMinutes()).padStart(2, '0');
    }
    // "HH:MM" 形式ならそのまま
    if (/^\d{1,2}:\d{2}$/.test(s)) return s;
    // "Sat Dec 30 1899 11:00:00 GMT+0900..." のような文字列から時刻を抽出
    var m = s.match(/(\d{1,2}:\d{2}):\d{2}/);
    if (m) return m[1];
    return s;
  };
  var shifts = rows.map(function(r) {
    return (r[2] || '') + '：' + formatTime(r[3]) + '〜' + formatTime(r[4]) + '（' + r[10] + 'h）';
  }).join('\n');
  // イベント名／実施業務: 重複を除去
  var titleSet = {};
  rows.forEach(function(r) {
    var t = (r[5] || '').trim();
    if (t) titleSet[t] = true;
  });
  var titles = Object.keys(titleSet).join('\n');
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

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// generateMonthlyReports: 指定月の全日分PDFを一括生成
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function generateMonthlyReports(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('「日報_2026」シートが見つかりません');

  const data = sheet.getDataRange().getValues();
  // 該当月のデータから日付の一覧を取得（重複排除・昇順）
  var dates = {};
  data.slice(1).forEach(function(row) {
    var d = toDateStr(row[1]);
    if (d && d.indexOf(yearMonth) === 0) {
      dates[d] = true;
    }
  });

  var dateList = Object.keys(dates).sort();
  if (dateList.length === 0) {
    throw new Error(yearMonth + ' のデータが見つかりません');
  }

  Logger.log(yearMonth + ': ' + dateList.length + '日分のPDFを生成します');

  var urls = [];
  dateList.forEach(function(dateStr) {
    var url = generateDailyReport(dateStr);
    urls.push(dateStr + ': ' + url);
  });

  Logger.log('一括生成完了:\n' + urls.join('\n'));
  return urls;
}

// ── 1ページ分のHTML ──────────────────────
function buildPage(d) {
  var nl2br = function(str) {
    return (str || '').replace(/\n/g, '<br>');
  };

  return ''
    + '<div class="page">'
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
    + '@page { size: A4 portrait; margin: 15mm 15mm 15mm 15mm; }'
    + 'body { font-family: "Noto Sans JP", "Hiragino Kaku Gothic Pro", "Yu Gothic", sans-serif; font-size: 10.5pt; color: #000; margin: 0; padding: 0; line-height: 1.25; }'
    + '.page { position: relative; }'
    + ''
    + 'h1 { text-align: center; font-size: 15pt; font-weight: bold; margin-bottom: 14px; }'
    + ''
    + '.header-row { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px; }'
    + '.date { font-size: 10.5pt; padding-top: 4px; }'
    + ''
    + '.stamp-table { border-collapse: collapse; margin-left: auto; }'
    + '.stamp-label { font-size: 9pt; text-align: left; padding: 2px 4px; border: none; }'
    + '.stamp-cell { width: 18mm; height: 18mm; border: 1px solid #000; }'
    + ''
    + '.report-table { width: 100%; border-collapse: collapse; margin-bottom: 12px; }'
    + '.report-table th, .report-table td { border: 1px solid #000; padding: 6px 10px; vertical-align: top; font-size: 10pt; line-height: 1.25; }'
    + '.report-table th { width: 22%; font-weight: bold; background: #f8f8f8; text-align: left; white-space: nowrap; }'
    + '.report-table td { width: 78%; }'
    + ''
    + '.company { text-align: center; font-size: 10.5pt; margin-top: 10px; }';
}
