// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 管理用スプレッドシート（CCBTテクニカル業務 実績管理 2026）の Apps Script
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 機能:
//   - メニュー「日報PDF → PDF出力」: 「設定」シート B1 の年月日でPDF生成
//                                    （要: 「テンプレート_日報」シート）
//   - メニュー「割り振り → 取り込み」: #2 のデータを「割り振り台帳」に増分同期
//   - メニュー「割り振り → 自動取り込みをON/OFF」: シート起動時の自動同期
//
// セットアップ:
//   1. 管理用スプレッドシートの「拡張機能 → Apps Script」にこのファイルを貼り付け
//   2. 「テンプレート_日報」シートを作成（PDF出力用）
//   3. メニューが現れない場合はシートを再読込
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ── 設定 ──────────────────────────
const CONFIG = {
  DATA_SPREADSHEET_ID: '16I5MK1Tqv2_UXi-I8VYHSO22iAO-joNpV_PrejHntVQ',  // 閲覧用（日報データ）
  SHEET_NAME: '日報_2026',
  TEMPLATE_SHEET_NAME: 'テンプレート_日報',
  DRIVE_FOLDER_ID: '1Nx9ALl1p1Riun68L9l9OJpG19UyCDpkj',
  RESPONSIBLE_PERSON: '伊藤友哉（arsaffix Inc.）',
  ALLOCATION_SHEET_NAME: '割り振り台帳',
};

// ── 4請求項目の定義（割り振り台帳用） ──────────────────
//   key: 内部キー
//   label: 列ヘッダー
//   col: 1-indexed の列番号
//   eligible: 計上可能なポスト区分
var BILLING_ITEMS = [
  { key: 'item1', label: '①実施計画 (h)',           col: 12, eligible: ['L'] },       // 実施計画策定・全体管理
  { key: 'item2', label: '②機器運用 (h)',           col: 13, eligible: ['L', 'S'] },  // 機器運用
  { key: 'item3', label: '③設営・技術支援 (h)',     col: 14, eligible: ['L', 'S'] },  // 設営管理・技術支援
  { key: 'item4', label: '④事業マネジメント (h)',   col: 15, eligible: ['L'] },       // 各事業におけるテクニカルマネジメント
];

// 割り振り台帳の列番号（1-indexed）
var ALLOC_COL = {
  TIMESTAMP: 1, DATE: 2, NAME: 3, POST: 4, START: 5, END: 6,
  HOURS: 7, EVENT: 8, TASKS: 9, CONTENT: 10, NOTES: 11,
  ITEM1: 12, ITEM2: 13, ITEM3: 14, ITEM4: 15,
  ALLOC_SUM: 16, ALLOC_DIFF: 17, MEMO: 18,
};

var ALLOCATION_HEADERS = [
  'タイムスタンプ', '日付', '氏名', 'ポスト', '開始', '終了',
  '勤務時間 (h)', 'イベント', '実施事項', '業務内容', '特記事項',
  '①実施計画 (h)', '②機器運用 (h)', '③設営・技術支援 (h)', '④事業マネジメント (h)',
  '配分計 (h)', '差異 (h)', 'メモ',
];

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// メニュー: スプレッドシートを開いたときにカスタムメニューを追加
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('日報PDF')
    .addItem('PDF出力（設定シートの年月を使用）', 'runFromSheet')
    .addToUi();
  ui.createMenu('割り振り')
    .addItem('#2 から最新データを取り込み', 'runSyncAllocation')
    .addSeparator()
    .addItem('自動取り込みをON（推奨）', 'installAutoSyncTrigger')
    .addItem('自動取り込みをOFF', 'removeAutoSyncTrigger')
    .addToUi();
}

function runFromSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('設定');

  if (!sheet) {
    sheet = ss.insertSheet('設定');
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
function generateDailyReport(dateStr) {
  if (/^\d{4}-\d{2}$/.test(dateStr)) {
    return generateMonthlyReports(dateStr);
  }

  var allRows = fetchAllRows_();
  var rows = allRows.filter(function(row) { return toDateStr(row[1]) === dateStr; });
  return generatePdfForDate_(dateStr, rows);
}

// ── データ取得（キャッシュ付き） ──────────────────
var cachedRows_ = null;
function fetchAllRows_() {
  if (cachedRows_) return cachedRows_;
  var dataSs = SpreadsheetApp.openById(CONFIG.DATA_SPREADSHEET_ID);
  var dataSheet = dataSs.getSheetByName(CONFIG.SHEET_NAME);
  if (!dataSheet) throw new Error('閲覧用スプレッドシートに「' + CONFIG.SHEET_NAME + '」シートが見つかりません');
  cachedRows_ = dataSheet.getDataRange().getValues().slice(1);
  return cachedRows_;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// generatePdfForDate_: 1日分のPDFを生成しDriveに保存
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function generatePdfForDate_(dateStr, rows) {
  if (rows.length === 0) {
    throw new Error(dateStr + ' のデータが見つかりません');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 日付を和暦に変換
  var date = new Date(dateStr + 'T00:00:00');
  var weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  var reiwa = date.getFullYear() - 2018;
  var dateDisplay = '令和' + reiwa + '年' + (date.getMonth() + 1) + '月' + date.getDate() + '日（' + weekdays[date.getDay()] + '）';

  // row: [タイムスタンプ, 日付, 氏名, ポスト区分, 開始, 終了, タイトル, 実施事項, 内容, 特記事項等, 気づき・振り返り, 勤務時間]
  //        0            1     2     3           4     5     6        7        8      9            10             11
  var formatTime = function(val) {
    if (!val) return '';
    var s = String(val);
    if (val instanceof Date) {
      return String(val.getHours()).padStart(2, '0') + ':' + String(val.getMinutes()).padStart(2, '0');
    }
    if (/^\d{1,2}:\d{2}$/.test(s)) return s;
    var m = s.match(/(\d{1,2}:\d{2}):\d{2}/);
    if (m) return m[1];
    return s;
  };
  var postLabel = function(v) { return v === 'L' ? 'リーダー' : v === 'S' ? 'サポーター' : ''; };
  var shifts = rows.map(function(r) {
    var post = r[3] ? '(' + postLabel(r[3]) + ')' : '';
    return (r[2] || '') + post + '：' + formatTime(r[4]) + '〜' + formatTime(r[5]) + '（' + r[11] + 'h）';
  }).join('\n');

  // イベント名／実施業務: 個々の項目単位で重複を除去
  var titleSet = {};
  rows.forEach(function(r) {
    var t = String(r[6] || '').trim();
    if (t) {
      t.split('\n').forEach(function(item) {
        var trimmed = item.trim();
        if (trimmed) titleSet[trimmed] = true;
      });
    }
  });
  var titles = Object.keys(titleSet).join('\n');

  // 実施事項・業務内容: 改行は「／」に置換して1行表示
  var combineField = function(idx) {
    return rows.map(function(r) {
      if (!r[idx]) return null;
      var text = String(r[idx]).replace(/\n/g, '／');
      if (rows.length > 1) {
        return '【' + (r[2] || '') + '】' + text;
      }
      return text;
    }).filter(Boolean).join('\n');
  };
  var tasks = combineField(7);
  var contents = combineField(8);

  // 特記事項等
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

  var replacements = {
    '{{date}}': dateDisplay,
    '{{title}}': titles,
    '{{tasks}}': tasks,
    '{{content}}': contents,
    '{{shift}}': shifts,
    '{{notes}}': notes || '特記事項なし',
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
          newVal = newVal.split(key).join(replacements[key] || '');
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

  try {
    var pdfBlob = fetchPdfWithRetry_(pdfUrl, fileName);
    var pdfFile = folder.createFile(pdfBlob);
    Logger.log('PDF保存完了: ' + pdfFile.getUrl());
    return pdfFile.getUrl();
  } finally {
    ss.deleteSheet(tmpSheet);
  }
}

// ── PDFエクスポート（429/5xx は指数バックオフでリトライ） ──
function fetchPdfWithRetry_(pdfUrl, fileName) {
  var maxAttempts = 5;
  var delayMs = 1500;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    var response = UrlFetchApp.fetch(pdfUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    var code = response.getResponseCode();
    if (code === 200) {
      return response.getBlob().setName(fileName + '.pdf');
    }
    if (code === 429 || code >= 500) {
      if (attempt === maxAttempts) {
        throw new Error('PDFエクスポートのレート制限に達しました（HTTP ' + code + '）。1〜2分ほど待ってから再実行してください。');
      }
      Utilities.sleep(delayMs);
      delayMs *= 2;
      continue;
    }
    throw new Error('PDFエクスポートに失敗しました（HTTP ' + code + '）');
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// generateMonthlyReports: 指定月の全日分PDFを一括生成
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function generateMonthlyReports(yearMonth) {
  var allRows = fetchAllRows_();

  // 日付ごとにグループ化
  var grouped = {};
  allRows.forEach(function(row) {
    var d = toDateStr(row[1]);
    if (d && d.substring(0, 7) === yearMonth) {
      if (!grouped[d]) grouped[d] = [];
      grouped[d].push(row);
    }
  });

  var dateList = Object.keys(grouped).sort();
  if (dateList.length === 0) {
    throw new Error(yearMonth + ' のデータが見つかりません');
  }

  Logger.log(yearMonth + ': ' + dateList.length + '日分のPDFを生成します');

  var urls = [];
  dateList.forEach(function(dateStr, idx) {
    if (idx > 0) Utilities.sleep(800);
    var url = generatePdfForDate_(dateStr, grouped[dateStr]);
    urls.push(dateStr + ': ' + url);
  });

  Logger.log('一括生成完了:\n' + urls.join('\n'));
  return urls;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 割り振り台帳: #2のデータを取り込み、4請求項目への時間配分を入力する場所
//   - 取り込みは増分のみ（タイムスタンプをキーに重複排除）
//   - 4項目列(L〜O)は0.25h刻みのデータ検証
//   - 配分計≠勤務時間 のとき差異列を赤
//   - Sポスト × 項目1/4 のとき該当セルを赤（制約違反警告）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ── メニュー: 手動取り込み ──
function runSyncAllocation() {
  var ui = SpreadsheetApp.getUi();
  try {
    var n = syncAllocationSheet();
    if (n === 0) {
      ui.alert('割り振り台帳: 新しいデータはありませんでした。');
    } else {
      ui.alert('割り振り台帳: ' + n + '件追加しました。');
    }
  } catch (err) {
    ui.alert('エラー: ' + err.message);
  }
}

// ── メニュー: 自動取り込みトリガーを有効化 ──
function installAutoSyncTrigger() {
  var ui = SpreadsheetApp.getUi();
  removeAutoSyncTriggers_();
  ScriptApp.newTrigger('onOpenAutoSync')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  ui.alert('自動取り込みを有効化しました。\n次回シートを開いたときから自動で取り込みます。');
}

function removeAutoSyncTrigger() {
  var n = removeAutoSyncTriggers_();
  SpreadsheetApp.getUi().alert('自動取り込みを解除しました（' + n + '件）。');
}

function removeAutoSyncTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'onOpenAutoSync') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}

// ── installable trigger ハンドラ（自動取り込み用） ──
function onOpenAutoSync() {
  try {
    syncAllocationSheet();
  } catch (e) {
    Logger.log('自動取り込み失敗: ' + e.message);
  }
}

// ── 割り振り台帳シートの初期化（無ければ作る） ──
function ensureAllocationSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ALLOCATION_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(CONFIG.ALLOCATION_SHEET_NAME);

  // ヘッダー
  sheet.getRange(1, 1, 1, ALLOCATION_HEADERS.length)
    .setValues([ALLOCATION_HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8eaed');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);

  // 列幅
  sheet.setColumnWidth(ALLOC_COL.TIMESTAMP, 140);
  sheet.setColumnWidth(ALLOC_COL.DATE, 90);
  sheet.setColumnWidth(ALLOC_COL.NAME, 80);
  sheet.setColumnWidth(ALLOC_COL.POST, 50);
  sheet.setColumnWidth(ALLOC_COL.START, 60);
  sheet.setColumnWidth(ALLOC_COL.END, 60);
  sheet.setColumnWidth(ALLOC_COL.HOURS, 80);
  sheet.setColumnWidth(ALLOC_COL.EVENT, 200);
  sheet.setColumnWidth(ALLOC_COL.TASKS, 200);
  sheet.setColumnWidth(ALLOC_COL.CONTENT, 200);
  sheet.setColumnWidth(ALLOC_COL.NOTES, 150);
  for (var c = ALLOC_COL.ITEM1; c <= ALLOC_COL.ITEM4; c++) sheet.setColumnWidth(c, 130);
  sheet.setColumnWidth(ALLOC_COL.ALLOC_SUM, 80);
  sheet.setColumnWidth(ALLOC_COL.ALLOC_DIFF, 80);
  sheet.setColumnWidth(ALLOC_COL.MEMO, 150);

  // データ検証: 4項目列を 0.25h刻み（≥0）に
  var maxRows = sheet.getMaxRows() - 1;
  BILLING_ITEMS.forEach(function(item) {
    var letter = columnLetter_(item.col);
    var rule = SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(ISBLANK(' + letter + '2),AND(' + letter + '2>=0,MOD(' + letter + '2*4,1)=0))')
      .setHelpText('0.25h刻みで入力してください（例: 0.25, 0.5, 0.75, 1.0）')
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, item.col, maxRows, 1).setDataValidation(rule);
  });

  // 条件付き書式
  var rules = sheet.getConditionalFormatRules();

  // 差異列が0以外（勤務時間入力ありかつ差異≠0）→ 赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($G2<>"",$Q2<>"",$Q2<>0)')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange(2, ALLOC_COL.ALLOC_DIFF, maxRows, 1)])
    .build());

  // Sポスト × 項目1（L列）→ 赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($D2="S",$L2<>"")')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange(2, ALLOC_COL.ITEM1, maxRows, 1)])
    .build());

  // Sポスト × 項目4（O列）→ 赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($D2="S",$O2<>"")')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange(2, ALLOC_COL.ITEM4, maxRows, 1)])
    .build());

  sheet.setConditionalFormatRules(rules);

  return sheet;
}

// ── #2 から最新データを取り込む（増分のみ、タイムスタンプで重複排除） ──
function syncAllocationSheet() {
  var sheet = ensureAllocationSheet_();
  var dataRows = fetchAllRows_();

  // 既存タイムスタンプ集合
  var existing = {};
  if (sheet.getLastRow() > 1) {
    var values = sheet.getRange(2, ALLOC_COL.TIMESTAMP, sheet.getLastRow() - 1, 1).getValues();
    values.forEach(function(row) {
      var ts = row[0];
      if (ts instanceof Date) existing[ts.getTime()] = true;
    });
  }

  // 新規行を抽出
  var newRows = [];
  dataRows.forEach(function(r) {
    var ts = r[0];
    if (!(ts instanceof Date)) return;
    if (existing[ts.getTime()]) return;

    newRows.push([
      ts,                                             // A タイムスタンプ
      toDateStr(r[1]),                                // B 日付
      r[2] || '',                                     // C 氏名
      r[3] || '',                                     // D ポスト
      formatTimeForDisplay_(r[4]),                    // E 開始
      formatTimeForDisplay_(r[5]),                    // F 終了
      r[11] === '' || r[11] == null ? '' : r[11],     // G 勤務時間
      String(r[6] || '').replace(/\n/g, ' / '),       // H イベント
      String(r[7] || '').replace(/\n/g, ' / '),       // I 実施事項
      String(r[8] || '').replace(/\n/g, ' / '),       // J 業務内容
      String(r[9] || '').replace(/\n/g, ' / '),       // K 特記事項
      '', '', '', '',                                 // L-O 4項目（空、手入力）
      '', '',                                         // P-Q 配分計・差異（数式は後置）
      '',                                             // R メモ
    ]);
  });

  if (newRows.length === 0) return 0;

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, newRows.length, ALLOCATION_HEADERS.length).setValues(newRows);

  // P列(配分計)・Q列(差異)に行ごとの数式
  var formulasP = newRows.map(function(_, i) {
    var r = startRow + i;
    return ['=IF(SUM(L' + r + ':O' + r + ')=0,"",SUM(L' + r + ':O' + r + '))'];
  });
  var formulasQ = newRows.map(function(_, i) {
    var r = startRow + i;
    return ['=IF(P' + r + '="","",G' + r + '-P' + r + ')'];
  });
  sheet.getRange(startRow, ALLOC_COL.ALLOC_SUM, newRows.length, 1).setFormulas(formulasP);
  sheet.getRange(startRow, ALLOC_COL.ALLOC_DIFF, newRows.length, 1).setFormulas(formulasQ);

  return newRows.length;
}

// ── 時刻表示フォーマット（HH:mm） ──
function formatTimeForDisplay_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
  }
  var s = String(val);
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  var m = s.match(/(\d{1,2}:\d{2}):\d{2}/);
  if (m) return m[1];
  return s;
}

// ── 1-indexed の列番号 → A1記法の列文字 ──
function columnLetter_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
