// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 管理用スプレッドシート（CCBTテクニカル業務 実績管理 2026）の Apps Script
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 機能:
//   - 日報PDF → 日次PDF出力             : 「設定」B1の年月日で日報PDFを生成
//   - 日報PDF → 月次サマリPDF出力       : カテゴリ別の月次サマリPDF（日数・延べ人数）
//   - 割り振り → #2 から最新データ取り込み : 日報→「割り振り台帳」へ増分同期
//   - 割り振り → 月次サマリPDF出力       : 4請求項目別の月次サマリPDF（クライアント送付用）
//   - 割り振り → 請求集計シートを更新   : 社内確認用シート（L/S別ポスト数・時間・金額の詳細）
//   - 割り振り → 自動取り込みON/OFF     : シート起動時の自動同期
//
// セットアップ:
//   1. 管理用スプレッドシートの「拡張機能 → Apps Script」にこのファイルを貼り付け
//   2. 以下のテンプレートシートを作成（プレースホルダ付き）:
//        ・「テンプレート_日報」       … 日次PDF用
//        ・「テンプレート_月次サマリ」 … カテゴリ別 月次サマリPDF用
//        ・「テンプレート_請求サマリ」 … 4請求項目別 月次サマリPDF用（クライアント送付）
//   3. メニューが現れない場合はシートを再読込
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ── 設定 ──────────────────────────
const CONFIG = {
  DATA_SPREADSHEET_ID: '16I5MK1Tqv2_UXi-I8VYHSO22iAO-joNpV_PrejHntVQ',  // 閲覧用（日報データ）
  SHEET_NAME: '日報_2026',
  TEMPLATE_SHEET_NAME: 'テンプレート_日報',
  SUMMARY_TEMPLATE_SHEET_NAME: 'テンプレート_月次サマリ',
  BILLING_SUMMARY_TEMPLATE_SHEET_NAME: 'テンプレート_請求サマリ',
  DRIVE_FOLDER_ID: '1Nx9ALl1p1Riun68L9l9OJpG19UyCDpkj',
  RESPONSIBLE_PERSON: '伊藤友哉（arsaffix Inc.）',
  ALLOCATION_SHEET_NAME: '割り振り台帳',
  ALLOCATION_SUMMARY_SHEET_NAME: '割り振り集計',
  BILLING_DETAIL_SHEET_NAME: '請求集計',
  SETTINGS_DAILY_SHEET_NAME: '設定_日報',
  SETTINGS_BILLING_SHEET_NAME: '設定_請求集計',
  // 請求単価（クライアント請求用 / 年間支払計画書ベース、社内ロジック）
  LEADER_DAILY_RATE: 36000,     // L: テクニカルディレクターおよびリーダー（円/日=ポスト）
  SUPPORTER_DAILY_RATE: 29400,  // S: テクニカルサポーター（円/日=ポスト）
  TAX_RATE: 0.10,               // 消費税
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

// ── カテゴリ別月次サマリのカテゴリ定義 ──────────────────
// label がフォーム送信値と一致する。データ内の未知カテゴリは「その他」に集約。
var SUMMARY_CATEGORIES = [
  { key: 'incubation', label: 'アートインキュベーションプログラム' },
  { key: 'camp',       label: 'キャンプ' },
  { key: 'showcase',   label: 'ショーケース' },
  { key: 'workshop',   label: 'ワークショップ' },
  { key: 'meetup',     label: 'ミートアップ' },
  { key: 'lab',        label: 'ラボ運営および施設管理' },
  { key: 'other',      label: 'その他' },
];

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// メニュー: スプレッドシートを開いたときにカスタムメニューを追加
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('日報PDF')
    .addItem('日次PDF出力（設定シートの年月を使用）', 'runFromSheet')
    .addItem('カテゴリ別 月次サマリPDF出力', 'runSummaryFromSheet')
    .addToUi();
  ui.createMenu('割り振り')
    .addItem('#2 から最新データを取り込み', 'runSyncAllocation')
    .addSeparator()
    .addItem('集計シートを作成/更新', 'runUpsertAllocationSummary')
    .addItem('請求集計シートを更新（社内確認用）', 'runUpsertBillingDetail')
    .addItem('請求項目別 月次サマリPDF出力（クライアント送付用）', 'runBillingSummaryFromSheet')
    .addSeparator()
    .addItem('自動取り込みをON（推奨）', 'installAutoSyncTrigger')
    .addItem('自動取り込みをOFF', 'removeAutoSyncTrigger')
    .addToUi();
  ui.createMenu('設定')
    .addItem('設定シートを準備/再構築（設定_日報 + 設定_請求集計）', 'runRebuildSettingsSheets')
    .addToUi();
}

function runFromSheet() {
  var ui = SpreadsheetApp.getUi();
  ensureDailySettingsSheet_();
  var dateStr = readDayOrMonthFromDailySettings_();

  if (!dateStr || (!/^\d{4}-\d{2}$/.test(dateStr) && !/^\d{4}-\d{2}-\d{2}$/.test(dateStr))) {
    ui.alert('「設定_日報」シートの B3 に YYYY-MM または YYYY-MM-DD を入力してください（例: 2026-04 または 2026-04-15）。');
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

// ── 月次サマリ用: 設定_日報 の B3 から対象月を取得して実行（YYYY-MM のときのみ有効） ──
function runSummaryFromSheet() {
  var ui = SpreadsheetApp.getUi();
  ensureDailySettingsSheet_();
  var ym = readDayOrMonthFromDailySettings_();

  if (!/^\d{4}-\d{2}$/.test(ym)) {
    ui.alert('「設定_日報」シートの B3 に対象月を YYYY-MM 形式で入力してください（例: 2026-04）。\n月次サマリは日指定では出力できません。');
    return;
  }

  ui.alert('月次サマリPDFを生成します: ' + ym);

  try {
    var url = generateMonthlySummary(ym);
    ui.alert('完了: 月次サマリPDFを出力しました。\n' + url);
  } catch (err) {
    ui.alert('エラー: ' + err.message);
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 設定シート（用途別に2枚に分離）
//
//   「設定_日報」     B3: 対象日(YYYY-MM-DD) or 対象月(YYYY-MM)
//                     → 日次PDF / カテゴリ別月次サマリPDF が使用
//
//   「設定_請求集計」 B3: 対象月(YYYY-MM)
//                     B5: L 前月繰越H
//                     B6: S 前月繰越H
//                     → 請求集計シート / 請求サマリPDF が使用
//
//   旧「設定」シートが残っている場合、初回アクセス時に値をベストエフォートで
//   新シートへ移植する（旧シートは消さない、ユーザーが手動で削除）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function ensureDailySettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = CONFIG.SETTINGS_DAILY_SHEET_NAME;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  // 旧シートから値を拾う（あれば）
  var seed = pickDayOrMonthFromLegacy_(ss);

  sheet = ss.insertSheet(name);
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 200);
  sheet.getRange('A1').setValue('設定_日報')
    .setFontWeight('bold').setFontSize(14).setBackground('#fff4d6');
  sheet.getRange('A2').setValue('↑ 日次PDF / カテゴリ別月次サマリPDF で使用')
    .setFontColor('#666666').setFontStyle('italic');
  sheet.getRange('A3').setValue('対象（YYYY-MM-DD or YYYY-MM）');
  sheet.getRange('B3').setBackground('#fff9e6').setNumberFormat('@')
    .setNote('YYYY-MM-DD で特定日の日報、YYYY-MM でその月の日報まとめ または 月次サマリ');
  if (seed) sheet.getRange('B3').setValue(seed);
  return sheet;
}

function ensureBillingSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = CONFIG.SETTINGS_BILLING_SHEET_NAME;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  // 旧シートから値を拾う（あれば）
  var seed = pickBillingFromLegacy_(ss);

  sheet = ss.insertSheet(name);
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 200);
  sheet.getRange('A1').setValue('設定_請求集計')
    .setFontWeight('bold').setFontSize(14).setBackground('#fff4d6');
  sheet.getRange('A2').setValue('↑ 請求集計シート更新 / 請求サマリPDF で使用')
    .setFontColor('#666666').setFontStyle('italic');
  sheet.getRange('A3').setValue('対象月（YYYY-MM）');
  sheet.getRange('B3').setBackground('#fff9e6').setNumberFormat('@')
    .setNote('YYYY-MM 形式で入力（例: 2026-04）');
  sheet.getRange('A5').setValue('L 前月繰越H');
  sheet.getRange('A6').setValue('S 前月繰越H');
  sheet.getRange('B5:B6').setBackground('#fff9e6').setNumberFormat('0.00');
  sheet.getRange('B5').setNote('前月から繰り越したL区分の時間。4月のみ前年度シートから、5月以降は前月「請求集計」の翌月繰越Hを転記');
  sheet.getRange('B6').setNote('前月から繰り越したS区分の時間。4月のみ前年度シートから、5月以降は前月「請求集計」の翌月繰越Hを転記');

  if (seed) {
    if (seed.month) sheet.getRange('B3').setValue(seed.month);
    if (typeof seed.L === 'number') sheet.getRange('B5').setValue(seed.L);
    if (typeof seed.S === 'number') sheet.getRange('B6').setValue(seed.S);
  }
  return sheet;
}

// 旧「設定」シートから日次PDF用の値を拾う（移植用）
function pickDayOrMonthFromLegacy_(ss) {
  var old = ss.getSheetByName('設定');
  if (!old) return '';
  // 新統合形式(B8) → 旧オリジナル形式(B1) の順で探す
  var b8 = String(old.getRange('B8').getValue() || '').trim();
  if (/^\d{4}-\d{2}(-\d{2})?$/.test(b8)) return b8;
  var b1 = String(old.getRange('B1').getValue() || '').trim();
  if (/^\d{4}-\d{2}(-\d{2})?$/.test(b1)) return b1;
  return '';
}

// 旧「設定」シートから請求用の値を拾う（移植用）
function pickBillingFromLegacy_(ss) {
  var old = ss.getSheetByName('設定');
  if (!old) return null;
  var result = {};
  // 月: 新統合形式 B4 → 旧 B1 (YYYY-MM のときのみ)
  var b4 = String(old.getRange('B4').getValue() || '').trim();
  if (/^\d{4}-\d{2}$/.test(b4)) {
    result.month = b4;
  } else {
    var b1 = String(old.getRange('B1').getValue() || '').trim();
    if (/^\d{4}-\d{2}$/.test(b1)) result.month = b1;
  }
  // 繰越: 新統合形式 B12/B13 → 旧 B2/B3
  var l = old.getRange('B12').getValue();
  if (!(typeof l === 'number')) l = old.getRange('B2').getValue();
  var s = old.getRange('B13').getValue();
  if (!(typeof s === 'number')) s = old.getRange('B3').getValue();
  if (typeof l === 'number' && l >= 0) result.L = l;
  if (typeof s === 'number' && s >= 0) result.S = s;
  return result;
}

function readDayOrMonthFromDailySettings_() {
  var sheet = ensureDailySettingsSheet_();
  return String(sheet.getRange('B3').getValue() || '').trim();
}

function readMonthFromBillingSettings_() {
  var sheet = ensureBillingSettingsSheet_();
  return String(sheet.getRange('B3').getValue() || '').trim();
}

function readCarryFromBillingSettings_() {
  var sheet = ensureBillingSettingsSheet_();
  var bL = sheet.getRange('B5').getValue();
  var bS = sheet.getRange('B6').getValue();
  var nL = typeof bL === 'number' ? bL : parseFloat(bL);
  var nS = typeof bS === 'number' ? bS : parseFloat(bS);
  return {
    L: (isFinite(nL) && nL >= 0) ? nL : 0,
    S: (isFinite(nS) && nS >= 0) ? nS : 0,
  };
}

function runRebuildSettingsSheets() {
  var ui = SpreadsheetApp.getUi();
  try {
    ensureDailySettingsSheet_();
    ensureBillingSettingsSheet_();
    ui.alert('「設定_日報」「設定_請求集計」シートを準備しました。'
      + '\n旧「設定」シートが残っている場合は手動で削除してください。');
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
      '', '',                                         // P-Q 配分計・差異(数式は後置)
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

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 割り振り集計シート: 月×4請求項目の合計をQUERYで動的に表示。
//   割り振り台帳を編集すると即反映される（数式のみ、スナップショットではない）。
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function runUpsertAllocationSummary() {
  var ui = SpreadsheetApp.getUi();
  try {
    ensureAllocationSummarySheet_();
    ui.alert('「' + CONFIG.ALLOCATION_SUMMARY_SHEET_NAME + '」シートを作成/更新しました。');
  } catch (err) {
    ui.alert('エラー: ' + err.message);
  }
}

function ensureAllocationSummarySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allocName = CONFIG.ALLOCATION_SHEET_NAME;
  if (!ss.getSheetByName(allocName)) {
    throw new Error('「' + allocName + '」シートが見つかりません。先に取り込みを実行してください。');
  }

  var name = CONFIG.ALLOCATION_SUMMARY_SHEET_NAME;
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();

  // 配列リテラルで [yyyy-mm, L, M, N, O, 合計] の6列を組み立て、QUERYで月ごとに合計。
  // 各数値列に N() を通すのは、L〜O に空文字列（手入力前の初期値）が混ざると
  // QUERY が文字列列と判定して AVG_SUM_ONLY_NUMERIC エラーを出すため。
  var qb = "'" + allocName + "'!B2:B";
  var ql = "'" + allocName + "'!L2:L";
  var qm = "'" + allocName + "'!M2:M";
  var qn = "'" + allocName + "'!N2:N";
  var qo = "'" + allocName + "'!O2:O";

  var formula =
    '=QUERY({' +
      'ARRAYFORMULA(IF(' + qb + '="","",TEXT(' + qb + ',"yyyy-mm"))),' +
      'ARRAYFORMULA(N(' + ql + ')),' +
      'ARRAYFORMULA(N(' + qm + ')),' +
      'ARRAYFORMULA(N(' + qn + ')),' +
      'ARRAYFORMULA(N(' + qo + ')),' +
      'ARRAYFORMULA(N(' + ql + ')+N(' + qm + ')+N(' + qn + ')+N(' + qo + '))' +
    '},' +
    '"SELECT Col1, SUM(Col2), SUM(Col3), SUM(Col4), SUM(Col5), SUM(Col6) ' +
    'WHERE Col1 <> \'\' ' +
    'GROUP BY Col1 ORDER BY Col1 ' +
    'LABEL Col1 \'月\', SUM(Col2) \'①実施計画 (h)\', SUM(Col3) \'②機器運用 (h)\', SUM(Col4) \'③設営・技術支援 (h)\', SUM(Col5) \'④事業マネジメント (h)\', SUM(Col6) \'合計 (h)\'",' +
    '0)';

  sheet.getRange('A1').setFormula(formula);
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 100);
  for (var c = 2; c <= 5; c++) sheet.setColumnWidth(c, 140);
  sheet.setColumnWidth(6, 110);

  var maxRows = sheet.getMaxRows();
  sheet.getRange(2, 2, maxRows - 1, 5).setNumberFormat('0.00');
  sheet.getRange(1, 1, 1, 6)
    .setFontWeight('bold')
    .setBackground('#e8eaed')
    .setHorizontalAlignment('center');

  return sheet;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 請求集計シート（社内確認用）: 「設定」B1の月を Cルールで再集計し、
//   項目別の 時間・L/S別ポスト数・金額・延べ日数・延べ人数 をシートに表示。
//   請求書は別途社内フォーマットで作成。本シートは数字の根拠表示用。
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function runUpsertBillingDetail() {
  var ui = SpreadsheetApp.getUi();
  ensureBillingSettingsSheet_();
  var ym = readMonthFromBillingSettings_();
  if (!/^\d{4}-\d{2}$/.test(ym)) {
    ui.alert('「設定_請求集計」シートの B3 に対象月を YYYY-MM 形式で入力してください（例: 2026-04）。');
    return;
  }
  var carry = readCarryFromBillingSettings_();
  try {
    upsertBillingDetailSheet(ym, carry.L, carry.S);
    ui.alert('「' + CONFIG.BILLING_DETAIL_SHEET_NAME + '」シートを更新しました: ' + ym
      + '\n（前月繰越: L=' + carry.L + 'H, S=' + carry.S + 'H）');
  } catch (err) {
    ui.alert('エラー: ' + err.message);
  }
}

function upsertBillingDetailSheet(yearMonth, carryInL, carryInS) {
  var data = aggregateBillingData_(yearMonth, carryInL, carryInS);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = CONFIG.BILLING_DETAIL_SHEET_NAME;
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();

  var itemLabels = [
    '①実施計画策定・全体管理',
    '②機器運用',
    '③設営管理・技術支援',
    '④各事業におけるテクニカルマネジメント',
  ];
  var L = CONFIG.LEADER_DAILY_RATE;
  var S = CONFIG.SUPPORTER_DAILY_RATE;
  var ov = data.overall;

  // タイトル
  sheet.getRange('A1').setValue('請求集計 (' + yearMonth + ')')
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue(
    '社内確認用。L/S別に「当月実働H + 前月繰越H → floor(/8) = みなしP, 余りは翌月繰越H」。'
    + ' L単価=' + L + '円/ポスト, S単価=' + S + '円/ポスト'
  );

  // ── L/S 精算サマリ ────────────────────────────
  sheet.getRange('A4').setValue('【L/S別 ポスト精算】').setFontWeight('bold').setBackground('#dbe9f7');
  var lsHeaders = ['区分', '当月実働H', '前月繰越H', '使用可能H', 'みなしP', '翌月繰越H', '単価(円)', '金額(円)'];
  sheet.getRange(5, 1, 1, lsHeaders.length).setValues([lsHeaders])
    .setFontWeight('bold').setBackground('#e8eaed').setHorizontalAlignment('center');
  var lsRows = [
    ['L', ov.hoursL, ov.carryInL, ov.availL, ov.deemedPostsL, ov.carryOutL, L, ov.deemedPostsL * L],
    ['S', ov.hoursS, ov.carryInS, ov.availS, ov.deemedPostsS, ov.carryOutS, S, ov.deemedPostsS * S],
  ];
  sheet.getRange(6, 1, lsRows.length, lsHeaders.length).setValues(lsRows);
  // 数値書式
  sheet.getRange(6, 2, 2, 3).setNumberFormat('0.00');  // 当月実働H / 前月繰越H / 使用可能H
  sheet.getRange(6, 5, 2, 1).setNumberFormat('0');     // みなしP 整数
  sheet.getRange(6, 6, 2, 1).setNumberFormat('0.00');  // 翌月繰越H
  sheet.getRange(6, 7, 2, 2).setNumberFormat('#,##0'); // 単価・金額

  // 税抜/消費税/税込
  sheet.getRange(9, 7).setValue('税抜小計').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(9, 8).setValue(ov.subtotal).setFontWeight('bold').setNumberFormat('#,##0');
  sheet.getRange(10, 7).setValue('消費税(' + Math.round(CONFIG.TAX_RATE * 100) + '%)').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(10, 8).setValue(ov.tax).setFontWeight('bold').setNumberFormat('#,##0');
  sheet.getRange(11, 7).setValue('税込合計').setFontWeight('bold').setBackground('#fff4d6').setHorizontalAlignment('right');
  sheet.getRange(11, 8).setValue(ov.grand).setFontWeight('bold').setBackground('#fff4d6').setNumberFormat('#,##0');

  // ── 翌月繰越 案内 ────────────────────────────
  sheet.getRange('A13').setValue(
    '★翌月繰越（次月の「設定」B2/B3 に転記）: L=' + ov.carryOutL + 'H, S=' + ov.carryOutS + 'H'
  ).setFontWeight('bold').setBackground('#fff4d6');

  // ── 項目別 内訳 ────────────────────────────
  sheet.getRange('A15').setValue('【項目別 内訳（みなしPを当月実働時間比で按分）】')
    .setFontWeight('bold').setBackground('#dbe9f7');
  var itemHeaders = ['項目', 'L時間(h)', 'L按分P', 'S時間(h)', 'S按分P', '金額計(円)', '延べ日数', '延べ人数'];
  sheet.getRange(16, 1, 1, itemHeaders.length).setValues([itemHeaders])
    .setFontWeight('bold').setBackground('#e8eaed').setHorizontalAlignment('center');

  var itemRows = data.items.map(function(it, i) {
    return [itemLabels[i], it.hoursL, it.postsL, it.hoursS, it.postsS, it.amount, it.days, it.personDays];
  });
  sheet.getRange(17, 1, itemRows.length, itemHeaders.length).setValues(itemRows);
  sheet.getRange(17, 2, itemRows.length, 1).setNumberFormat('0.00');    // L時間
  sheet.getRange(17, 3, itemRows.length, 1).setNumberFormat('0.0000');  // L按分P（小数4桁）
  sheet.getRange(17, 4, itemRows.length, 1).setNumberFormat('0.00');    // S時間
  sheet.getRange(17, 5, itemRows.length, 1).setNumberFormat('0.0000');  // S按分P（小数4桁）
  sheet.getRange(17, 6, itemRows.length, 1).setNumberFormat('#,##0');   // 金額
  sheet.getRange(17, 7, itemRows.length, 2).setNumberFormat('0');       // 延べ日数・人数

  // 更新日時
  var noteRow = 17 + itemRows.length + 2;
  sheet.getRange(noteRow, 1).setValue(
    '更新日時: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  );

  // 列幅
  sheet.setColumnWidth(1, 260);
  for (var c = 2; c <= 6; c++) sheet.setColumnWidth(c, 100);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 130);

  sheet.setFrozenRows(5);
  return data;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// aggregateBillingData_: 割り振り台帳から1ヶ月分を集計し、項目別の
//   時間/日数/延べ人数/ポスト数(L/S別)/金額を返す内部ヘルパー。
//
//   ポスト数の数え方ルール（"Cルール: 時間按分"）:
//     ・1行（=1人の1日分）= 1ポストとして扱う（勤務時間に関わらず）
//     ・項目内訳は、その日の項目時間比で按分
//       例: 4/15 佐藤さん(S) [②4h + ③2h] → ②=0.67ポスト, ③=0.33ポスト
//     ・L単価 36,000円/ポスト、S単価 29,400円/ポスト で金額算出
//     ・ポスト種別はD列。L/S以外の値はLとみなす（保守側）。
//
//   日数・延べ人数（既存仕様、互換維持）:
//     ・日数       … その項目に時間が入ったユニーク日付の数
//     ・延べ人数   … その項目に時間が入った (氏名×日付) のユニーク組み合わせ数
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function aggregateBillingData_(yearMonth, carryInL, carryInS) {
  if (!/^\d{4}-\d{2}$/.test(yearMonth)) {
    throw new Error('月次集計は YYYY-MM 形式で指定してください');
  }
  carryInL = (typeof carryInL === 'number' && isFinite(carryInL) && carryInL >= 0) ? carryInL : 0;
  carryInS = (typeof carryInS === 'number' && isFinite(carryInS) && carryInS >= 0) ? carryInS : 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allocSheet = ss.getSheetByName(CONFIG.ALLOCATION_SHEET_NAME);
  if (!allocSheet) {
    throw new Error('「' + CONFIG.ALLOCATION_SHEET_NAME + '」シートが見つかりません。先に取り込みを実行してください。');
  }

  var lastRow = allocSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('割り振り台帳にデータがありません');
  }

  var dates = allocSheet.getRange(2, ALLOC_COL.DATE, lastRow - 1, 1).getValues();
  var names = allocSheet.getRange(2, ALLOC_COL.NAME, lastRow - 1, 1).getValues();
  var posts = allocSheet.getRange(2, ALLOC_COL.POST, lastRow - 1, 1).getValues();
  var items = allocSheet.getRange(2, ALLOC_COL.ITEM1, lastRow - 1, 4).getValues();

  // 項目別 L/S 別 時間を集計
  var hoursL = [0, 0, 0, 0];          // L が項目kに投入した時間
  var hoursS = [0, 0, 0, 0];          // S が項目kに投入した時間
  var itemDates = [{}, {}, {}, {}];
  var itemPersonDays = [{}, {}, {}, {}];
  var allDates = {};
  var allPersonDays = {};
  var matched = 0;

  for (var i = 0; i < dates.length; i++) {
    var d = toDateStr(dates[i][0]);
    if (!d || d.substring(0, 7) !== yearMonth) continue;
    matched++;
    var name = String(names[i][0] || '').trim();
    var post = String(posts[i][0] || '').trim().toUpperCase();
    if (post !== 'L' && post !== 'S') post = 'L'; // 不明値は保守的にLへ

    var hasAny = false;
    for (var j = 0; j < 4; j++) {
      var hv = items[i][j];
      if (typeof hv !== 'number' || hv <= 0) continue;
      if (post === 'L') hoursL[j] += hv;
      else hoursS[j] += hv;
      itemDates[j][d] = true;
      if (name) itemPersonDays[j][name + '|' + d] = true;
      hasAny = true;
    }
    if (hasAny) {
      allDates[d] = true;
      if (name) allPersonDays[name + '|' + d] = true;
    }
  }

  if (matched === 0) {
    throw new Error(yearMonth + ' の割り振りデータが見つかりません');
  }

  // L/S 別の総時間 → 繰越を加味して みなしP と 翌月繰越 を算出
  var totalHoursL = 0, totalHoursS = 0;
  for (var k = 0; k < 4; k++) { totalHoursL += hoursL[k]; totalHoursS += hoursS[k]; }

  var availL = totalHoursL + carryInL;
  var availS = totalHoursS + carryInS;
  var deemedL = Math.floor(availL / 8);
  var deemedS = Math.floor(availS / 8);
  var carryOutL = Math.round((availL - deemedL * 8) * 100) / 100;
  var carryOutS = Math.round((availS - deemedS * 8) * 100) / 100;

  // 項目別配賦: みなしP を 当月実働の時間比で按分（繰越は項目に紐付けないため当月実働ベース）
  var L = CONFIG.LEADER_DAILY_RATE;
  var S = CONFIG.SUPPORTER_DAILY_RATE;
  var itemsOut = [];
  for (var m = 0; m < 4; m++) {
    var pL = totalHoursL > 0 ? (hoursL[m] / totalHoursL) * deemedL : 0;
    var pS = totalHoursS > 0 ? (hoursS[m] / totalHoursS) * deemedS : 0;
    var amount = pL * L + pS * S;
    itemsOut.push({
      key: BILLING_ITEMS[m].key,
      hoursL: hoursL[m],
      hoursS: hoursS[m],
      hours: hoursL[m] + hoursS[m],
      postsL: pL,
      postsS: pS,
      posts: pL + pS,
      amount: amount,
      days: Object.keys(itemDates[m]).length,
      personDays: Object.keys(itemPersonDays[m]).length,
    });
  }

  var subtotal = deemedL * L + deemedS * S;
  var tax = Math.round(subtotal * CONFIG.TAX_RATE);
  var grand = subtotal + tax;

  return {
    yearMonth: yearMonth,
    matched: matched,
    items: itemsOut,
    overall: {
      hoursL: totalHoursL,
      hoursS: totalHoursS,
      hours: totalHoursL + totalHoursS,
      carryInL: carryInL,
      carryInS: carryInS,
      availL: availL,
      availS: availS,
      deemedPostsL: deemedL,
      deemedPostsS: deemedS,
      posts: deemedL + deemedS,
      carryOutL: carryOutL,
      carryOutS: carryOutS,
      subtotal: subtotal,
      tax: tax,
      grand: grand,
      days: Object.keys(allDates).length,
      personDays: Object.keys(allPersonDays).length,
    },
  };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 4請求項目別 月次サマリPDF（クライアント送付用）: 割り振り台帳を月で集計し、
//   「テンプレート_請求サマリ」のプレースホルダを置換してPDF出力。
//   L/S別の「当月実働H + 前月繰越H → みなしP → 翌月繰越H」方式で算出。
//
//   プレースホルダ（テンプレート側で必要な分だけ配置）:
//     {{year_month}}                           … 例 "2026-04"
//     {{responsible}}                          … 担当者名
//     [L/S精算]
//     {{hours_l}} / {{hours_s}}                … 当月実働H（割り振り台帳ベース）
//     {{carry_in_l}} / {{carry_in_s}}          … 前月繰越H（設定B2/B3より）
//     {{avail_l}} / {{avail_s}}                … 使用可能H（実働+繰越）
//     {{deemed_l}} / {{deemed_s}}              … みなしP（=floor(使用可能/8)、整数）
//     {{carry_out_l}} / {{carry_out_s}}        … 翌月繰越H
//     [項目別 内訳]
//     {{item1_total}} 〜 {{item4_total}}       … 各項目の月時間合計（L+S、小数2桁）
//     {{item1_hours_l}} 〜 {{item4_hours_l}}   … 各項目のL時間（h）
//     {{item1_hours_s}} 〜 {{item4_hours_s}}   … 各項目のS時間（h）
//     {{item1_posts_l}} 〜 {{item4_posts_l}}   … 各項目のL按分P（小数2桁）
//     {{item1_posts_s}} 〜 {{item4_posts_s}}   … 各項目のS按分P（小数2桁）
//     {{item1_posts}}   〜 {{item4_posts}}     … 各項目のポスト合計（L+S、小数2桁）
//     {{item1_amount}}  〜 {{item4_amount}}    … 各項目の金額（円、整数）
//     {{item1_days}}  〜 {{item4_days}}        … 各項目の延べ日数（整数）
//     {{item1_persons}} 〜 {{item4_persons}}   … 各項目の延べ人数（整数）
//     [全体]
//     {{total_days}}                           … 全体のユニーク日数
//     {{total_persons}}                        … 全体の延べ人数
//     {{total_posts}}                          … 全体の みなしP合計
//     {{grand_total}}                          … 4項目の時間総合計（h、小数2桁）
//     {{subtotal_amount}}                      … 税抜小計（円）
//     {{tax_amount}}                           … 消費税額（円）
//     {{grand_amount}}                         … 税込合計（円）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function runBillingSummaryFromSheet() {
  var ui = SpreadsheetApp.getUi();
  ensureBillingSettingsSheet_();
  var ym = readMonthFromBillingSettings_();
  if (!/^\d{4}-\d{2}$/.test(ym)) {
    ui.alert('「設定_請求集計」シートの B3 に対象月を YYYY-MM 形式で入力してください（例: 2026-04）。');
    return;
  }

  var carry = readCarryFromBillingSettings_();
  ui.alert('請求項目別 月次サマリPDFを生成します: ' + ym
    + '\n（前月繰越: L=' + carry.L + 'H, S=' + carry.S + 'H）');

  try {
    var result = generateBillingSummary(ym, carry.L, carry.S);
    ui.alert('完了: PDFを出力しました。\n' + result.url);
  } catch (err) {
    ui.alert('エラー: ' + err.message);
  }
}

function generateBillingSummary(yearMonth, carryInL, carryInS) {
  var data = aggregateBillingData_(yearMonth, carryInL, carryInS);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName(CONFIG.BILLING_SUMMARY_TEMPLATE_SHEET_NAME);
  if (!templateSheet) {
    throw new Error('「' + CONFIG.BILLING_SUMMARY_TEMPLATE_SHEET_NAME + '」シートが見つかりません');
  }

  var fmt = function(n) { return (Math.round(n * 100) / 100).toString(); };           // 時間: 小数2桁
  var fmtPosts = function(n) { return Number(n).toFixed(4); };                          // ポスト: 小数4桁
  var fmtAmount = function(n) { return String(Math.round(n)); };
  var its = data.items;
  var ov = data.overall;

  var replacements = {
    '{{year_month}}': yearMonth,
    '{{responsible}}': CONFIG.RESPONSIBLE_PERSON,
    // L/S精算サマリ
    '{{hours_l}}': fmt(ov.hoursL),
    '{{hours_s}}': fmt(ov.hoursS),
    '{{carry_in_l}}': fmt(ov.carryInL),
    '{{carry_in_s}}': fmt(ov.carryInS),
    '{{avail_l}}': fmt(ov.availL),
    '{{avail_s}}': fmt(ov.availS),
    '{{deemed_l}}': String(ov.deemedPostsL),
    '{{deemed_s}}': String(ov.deemedPostsS),
    '{{carry_out_l}}': fmt(ov.carryOutL),
    '{{carry_out_s}}': fmt(ov.carryOutS),
    // 項目別 時間（L+S）
    '{{item1_total}}': fmt(its[0].hours),
    '{{item2_total}}': fmt(its[1].hours),
    '{{item3_total}}': fmt(its[2].hours),
    '{{item4_total}}': fmt(its[3].hours),
    '{{grand_total}}': fmt(ov.hours),
    // 項目別 L時間 / S時間
    '{{item1_hours_l}}': fmt(its[0].hoursL),
    '{{item2_hours_l}}': fmt(its[1].hoursL),
    '{{item3_hours_l}}': fmt(its[2].hoursL),
    '{{item4_hours_l}}': fmt(its[3].hoursL),
    '{{item1_hours_s}}': fmt(its[0].hoursS),
    '{{item2_hours_s}}': fmt(its[1].hoursS),
    '{{item3_hours_s}}': fmt(its[2].hoursS),
    '{{item4_hours_s}}': fmt(its[3].hoursS),
    // 項目別 延べ日数・延べ人数（説明資料用）
    '{{item1_days}}': String(its[0].days),
    '{{item2_days}}': String(its[1].days),
    '{{item3_days}}': String(its[2].days),
    '{{item4_days}}': String(its[3].days),
    '{{total_days}}': String(ov.days),
    '{{item1_persons}}': String(its[0].personDays),
    '{{item2_persons}}': String(its[1].personDays),
    '{{item3_persons}}': String(its[2].personDays),
    '{{item4_persons}}': String(its[3].personDays),
    '{{total_persons}}': String(ov.personDays),
    // 項目別 L/S 按分P（小数4桁）
    '{{item1_posts_l}}': fmtPosts(its[0].postsL),
    '{{item2_posts_l}}': fmtPosts(its[1].postsL),
    '{{item3_posts_l}}': fmtPosts(its[2].postsL),
    '{{item4_posts_l}}': fmtPosts(its[3].postsL),
    '{{item1_posts_s}}': fmtPosts(its[0].postsS),
    '{{item2_posts_s}}': fmtPosts(its[1].postsS),
    '{{item3_posts_s}}': fmtPosts(its[2].postsS),
    '{{item4_posts_s}}': fmtPosts(its[3].postsS),
    '{{item1_posts}}': fmtPosts(its[0].posts),
    '{{item2_posts}}': fmtPosts(its[1].posts),
    '{{item3_posts}}': fmtPosts(its[2].posts),
    '{{item4_posts}}': fmtPosts(its[3].posts),
    '{{total_posts}}': String(ov.posts),
    // 金額
    '{{item1_amount}}': fmtAmount(its[0].amount),
    '{{item2_amount}}': fmtAmount(its[1].amount),
    '{{item3_amount}}': fmtAmount(its[2].amount),
    '{{item4_amount}}': fmtAmount(its[3].amount),
    '{{subtotal_amount}}': fmtAmount(ov.subtotal),
    '{{tax_amount}}': fmtAmount(ov.tax),
    '{{grand_amount}}': fmtAmount(ov.grand),
  };

  var url = generateBillingPdf_(templateSheet, replacements, '請求サマリ_' + yearMonth);
  return { url: url, data: data };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// generateBillingPdf_: テンプレートをコピーしてプレースホルダ置換 → PDF化 → 保存。
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function generateBillingPdf_(templateSheet, replacements, fileName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmpName = '_tmp_' + fileName;
  var tmpSheet = templateSheet.copyTo(ss).setName(tmpName);

  var range = tmpSheet.getDataRange();
  var values = range.getValues();
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var cell = values[r][c];
      if (typeof cell !== 'string' || cell.indexOf('{{') === -1) continue;
      var newVal = cell;
      for (var key in replacements) {
        newVal = newVal.split(key).join(replacements[key]);
      }
      if (newVal !== cell) tmpSheet.getRange(r + 1, c + 1).setValue(newVal);
    }
  }

  SpreadsheetApp.flush();

  var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  var existing = folder.getFilesByName(fileName + '.pdf');
  while (existing.hasNext()) existing.next().setTrashed(true);

  var pdfUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?'
    + 'format=pdf'
    + '&gid=' + tmpSheet.getSheetId()
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
    Logger.log(fileName + ' PDF保存完了: ' + pdfFile.getUrl());
    return pdfFile.getUrl();
  } finally {
    ss.deleteSheet(tmpSheet);
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// generateMonthlySummary: カテゴリ別 月次サマリPDFを生成
//   集計仕様:
//     ・日数       … その実施業務がその月に報告されたユニーク日付の数
//     ・延べ人数   … (氏名 × 日付) のユニーク組み合わせ数
//                    （同じ人が10日関わったら10としてカウント）
//   ポスト区分（L/S）は区別しない。
//   データ内のカテゴリで定義に無いものは「その他」へ集約する。
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function generateMonthlySummary(yearMonth) {
  if (!/^\d{4}-\d{2}$/.test(yearMonth)) {
    throw new Error('月次サマリは YYYY-MM 形式で指定してください');
  }

  var allRows = fetchAllRows_();
  var monthRows = allRows.filter(function(row) {
    var d = toDateStr(row[1]);
    return d && d.substring(0, 7) === yearMonth;
  });

  if (monthRows.length === 0) {
    throw new Error(yearMonth + ' のデータが見つかりません');
  }

  // カテゴリ別集計用バケットを初期化
  var aggregates = {};
  SUMMARY_CATEGORIES.forEach(function(cat) {
    aggregates[cat.label] = { dates: {}, personDays: {} };
  });

  // 月全体の集計（カテゴリ問わず）
  var allDates = {};
  var allPersonDays = {};
  var allPersons = {};

  monthRows.forEach(function(row) {
    var dateStr = toDateStr(row[1]);
    var name = String(row[2] || '').trim();
    if (!dateStr) return;

    allDates[dateStr] = true;
    if (name) {
      allPersonDays[name + '|' + dateStr] = true;
      allPersons[name] = true;
    }

    var rawCats = String(row[6] || '').trim();
    if (!rawCats) return;

    rawCats.split('\n').forEach(function(item) {
      var cat = item.trim();
      if (!cat) return;
      var label = aggregates[cat] ? cat : 'その他';
      aggregates[label].dates[dateStr] = true;
      if (name) aggregates[label].personDays[name + '|' + dateStr] = true;
    });
  });

  // 表示用 yearMonth（令和○年○月）
  var ymParts = yearMonth.split('-');
  var year = parseInt(ymParts[0], 10);
  var month = parseInt(ymParts[1], 10);
  var reiwa = year - 2018;
  var yearMonthDisplay = '令和' + reiwa + '年' + month + '月';

  var replacements = {
    '{{yearMonth}}': yearMonthDisplay,
    '{{totalDays}}': Object.keys(allDates).length,
    '{{totalPersonDays}}': Object.keys(allPersonDays).length,
    '{{totalPersons}}': Object.keys(allPersons).length,
    '{{responsible}}': CONFIG.RESPONSIBLE_PERSON,
  };

  SUMMARY_CATEGORIES.forEach(function(cat) {
    var agg = aggregates[cat.label];
    replacements['{{days_' + cat.key + '}}'] = Object.keys(agg.dates).length;
    replacements['{{persons_' + cat.key + '}}'] = Object.keys(agg.personDays).length;
  });

  // テンプレートをコピーして置換 → PDFエクスポート
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName(CONFIG.SUMMARY_TEMPLATE_SHEET_NAME);
  if (!templateSheet) {
    throw new Error('「' + CONFIG.SUMMARY_TEMPLATE_SHEET_NAME + '」シートが見つかりません');
  }

  var tmpName = '_tmp_月次サマリ_' + yearMonth;
  var tmpSheet = templateSheet.copyTo(ss).setName(tmpName);

  var range = tmpSheet.getDataRange();
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cell = values[i][j];
      if (typeof cell === 'string' && cell.indexOf('{{') !== -1) {
        var newVal = cell;
        for (var key in replacements) {
          newVal = newVal.split(key).join(String(replacements[key]));
        }
        if (newVal !== cell) {
          tmpSheet.getRange(i + 1, j + 1).setValue(newVal);
        }
      }
    }
  }

  SpreadsheetApp.flush();

  var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  var fileName = '業務日報_月次サマリ_' + yearMonth;

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
    Logger.log('月次サマリPDF保存完了: ' + pdfFile.getUrl());
    return pdfFile.getUrl();
  } finally {
    ss.deleteSheet(tmpSheet);
  }
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
