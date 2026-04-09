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
  TEMPLATE_SHEET_NAME: 'テンプレート_日報',
  DRIVE_FOLDER_ID: '1nqFYO3mjOcwupn--p1Q7Bh8y-SAtTOl5',   // ← Google DriveフォルダIDに変更
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
  // row: [タイムスタンプ, 日付, 氏名, ポスト区分, 開始, 終了, タイトル, 実施事項, 内容, 特記事項等, 気づき・振り返り, 勤務時間]
  //        0            1     2     3           4     5     6        7        8      9            10             11
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
  var postLabel = function(v) { return v === 'L' ? 'リーダー' : v === 'S' ? 'サポーター' : ''; };
  var shifts = rows.map(function(r) {
    var post = r[3] ? '(' + postLabel(r[3]) + ')' : '';
    return (r[2] || '') + post + '：' + formatTime(r[4]) + '〜' + formatTime(r[5]) + '（' + r[11] + 'h）';
  }).join('\n');
  // イベント名／実施業務: 重複を除去
  var titleSet = {};
  rows.forEach(function(r) {
    var t = (r[6] || '').trim();
    if (t) titleSet[t] = true;
  });
  var titles = Object.keys(titleSet).join('\n');
  // 実施事項・業務内容: 【名前】内容... を同一行に（コンパクト表示）
  var combineField = function(idx) {
    return rows.map(function(r) {
      if (!r[idx]) return null;
      if (rows.length > 1) {
        return '【' + (r[2] || '') + '】' + r[idx];
      }
      return String(r[idx]);
    }).filter(Boolean).join('\n');
  };
  var tasks = combineField(7);
  var contents = combineField(8);
  // 特記事項等: 名前と内容を同一行に（【名前】内容...）
  var notes = rows.map(function(r) {
    if (!r[9]) return null;
    if (rows.length > 1) {
      return '【' + (r[2] || '') + '】' + r[9];
    }
    return String(r[9]);
  }).filter(Boolean).join('\n');

  // ── テンプレートシートをコピーしてデータを流し込む ──
  var templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET_NAME);
  if (!templateSheet) throw new Error('「' + CONFIG.TEMPLATE_SHEET_NAME + '」シートが見つかりません');

  var tmpName = '_tmp_日報_' + dateStr;
  var tmpSheet = templateSheet.copyTo(ss).setName(tmpName);

  // プレースホルダを置換
  var replacements = {
    '{{date}}': dateDisplay,
    '{{title}}': titles,
    '{{tasks}}': tasks,
    '{{content}}': contents,
    '{{shift}}': shifts,
    '{{notes}}': notes,
    '{{responsible}}': CONFIG.RESPONSIBLE_PERSON,
  };

  var range = tmpSheet.getDataRange();
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cell = values[i][j];
      if (typeof cell === 'string' && cell.indexOf('{{') !== -1) {
        var newVal = cell;
        for (var key in replacements) {
          newVal = newVal.replace(key, replacements[key] || '');
        }
        if (newVal !== cell) {
          tmpSheet.getRange(i + 1, j + 1).setValue(newVal);
        }
      }
    }
  }

  SpreadsheetApp.flush();

  // ── 一時シートをPDFとしてエクスポート ──
  var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  var fileName = '業務日報_' + dateStr;

  // 既存の同名PDFを削除
  var existing = folder.getFilesByName(fileName + '.pdf');
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  var ssId = ss.getId();
  var sheetId = tmpSheet.getSheetId();
  var pdfUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?'
    + 'format=pdf'
    + '&gid=' + sheetId
    + '&size=A4'
    + '&portrait=true'
    + '&fitw=true'
    + '&gridlines=false'
    + '&printtitle=false'
    + '&sheetnames=false'
    + '&pagenum=UNDEFINED'
    + '&fzr=false';

  var pdfBlob = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(fileName + '.pdf');

  var pdfFile = folder.createFile(pdfBlob);

  // 一時シートを削除
  ss.deleteSheet(tmpSheet);

  Logger.log('PDF保存完了: ' + pdfFile.getUrl());
  return pdfFile.getUrl();
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

