// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// テクニカルサポート業務日報 — Google Apps Script（閲覧用）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 使い方:
//   1. Googleスプレッドシートの「拡張機能 → Apps Script」に貼り付け
//   2. デプロイ → ウェブアプリ（実行:自分, アクセス:全員）
//   3. メニュー「日報メンテナンス → 時刻編集時の勤務時間 自動再計算をON」を一度実行
//   ※ PDF生成は管理用スプレッドシート（gas_code_admin.js）で行う
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ── 設定 ──────────────────────────
var CONFIG = {
  SHEET_NAME: '日報_2026',
};

// 日報シートの列番号（1-indexed）
var COL = {
  TIMESTAMP: 1, DATE: 2, NAME: 3, POST: 4, START: 5, END: 6,
  TITLE: 7, TASKS: 8, CONTENT: 9, NOTES: 10, REFLECTION: 11,
  HOURS: 12, WORKTYPE: 13,
};

// ── 数式インジェクション防止 ──────────────────
function sanitize(val) {
  if (typeof val !== 'string') return val;
  return /^[=+\-@]/.test(val) ? "'" + val : val;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 勤務時間の算出
//   L列（勤務時間）は数式ではなく値なので、時刻を直したら必ず再計算が要る。
//   doPost と onEditRecalcHours の両方がこの関数を通る。
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// 'HH:mm' / 'HH:mm:ss' / Date を 0時からの分に変換（読めなければ null）
function toMinutes_(val) {
  if (val === '' || val == null) return null;
  if (val instanceof Date) return val.getHours() * 60 + val.getMinutes();
  var m = /^(\d{1,2}):(\d{2})/.exec(String(val).trim());
  return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : null;
}

// 終了が開始より前なら日跨ぎとみなして24時間を足す
function calcWorkHours_(startVal, endVal) {
  var s = toMinutes_(startVal);
  var e = toMinutes_(endVal);
  if (s === null || e === null) return '';
  var mins = e - s;
  if (mins < 0) mins += 24 * 60;
  return mins / 60;
}

function formatDateCell_(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(val == null ? '' : val);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// doPost: フォームからのデータ受信
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
      sheet.appendRow([
        'タイムスタンプ', '日付', '氏名', 'ポスト区分',
        '開始時間', '終了時間',
        'イベント名／実施業務', '実施事項', '業務内容', '特記事項等',
        '気づき・振り返り', '勤務時間', '勤務形態'
      ]);
    }

    sheet.appendRow([
      new Date(),
      sanitize(data.date),
      sanitize(data.member),
      sanitize(data.post || ''),
      sanitize(data.start_time),
      sanitize(data.end_time),
      sanitize(data.title),
      sanitize(data.tasks),
      sanitize(data.content),
      sanitize(data.notes),
      sanitize(data.reflection),
      calcWorkHours_(data.start_time, data.end_time),
      sanitize(data.worktype || ''),  // 勤務形態: 現地/リモート（arsaffix内部用・PDF非掲載）
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
// メニュー
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function onOpen() {
  SpreadsheetApp.getUi().createMenu('日報メンテナンス')
    .addItem('時刻編集時の勤務時間 自動再計算をON（推奨）', 'installRecalcTrigger')
    .addItem('自動再計算をOFF', 'removeRecalcTrigger')
    .addSeparator()
    .addItem('勤務時間を一括で再計算', 'runRecalcAllHours')
    .addToUi();
}

// ── 開始・終了を編集したら勤務時間を再計算（installable な onEdit トリガー） ──
function onEditRecalcHours(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  // 編集範囲が E列(開始)〜F列(終了) にかかっているときだけ動く
  var firstCol = e.range.getColumn();
  var lastCol = firstCol + e.range.getNumColumns() - 1;
  if (lastCol < COL.START || firstCol > COL.END) return;

  var firstRow = Math.max(e.range.getRow(), 2);          // ヘッダー行は対象外
  var lastRow = e.range.getRow() + e.range.getNumRows() - 1;
  if (lastRow < firstRow) return;

  var n = lastRow - firstRow + 1;
  var times = sheet.getRange(firstRow, COL.START, n, 2).getValues();
  var hours = sheet.getRange(firstRow, COL.HOURS, n, 1).getValues();
  var next = times.map(function(t, i) {
    var h = calcWorkHours_(t[0], t[1]);
    return [h === '' ? hours[i][0] : h];                 // 時刻が読めない行は触らない
  });
  sheet.getRange(firstRow, COL.HOURS, n, 1).setValues(next);
}

function installRecalcTrigger() {
  var ui = SpreadsheetApp.getUi();
  removeRecalcTriggers_();
  ScriptApp.newTrigger('onEditRecalcHours')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  ui.alert('自動再計算を有効化しました。\n以降、開始・終了を直すと勤務時間が自動で計算し直されます。');
}

function removeRecalcTrigger() {
  var n = removeRecalcTriggers_();
  SpreadsheetApp.getUi().alert('自動再計算を解除しました（' + n + '件）。');
}

function removeRecalcTriggers_() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditRecalcHours') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}

// ── メニュー: 全行の勤務時間を開始・終了から再計算（確認ダイアログあり） ──
function runRecalcAllHours() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { ui.alert('「' + CONFIG.SHEET_NAME + '」シートが見つかりません。'); return; }

  var last = sheet.getLastRow();
  if (last < 2) { ui.alert('データがありません。'); return; }

  var n = last - 1;
  var dates = sheet.getRange(2, COL.DATE, n, 1).getValues();
  var names = sheet.getRange(2, COL.NAME, n, 1).getValues();
  var times = sheet.getRange(2, COL.START, n, 2).getValues();
  var cur   = sheet.getRange(2, COL.HOURS, n, 1).getValues();

  var changed = [];   // [シート行番号, 新しい勤務時間]
  var diffs = [];
  var unreadable = 0;
  for (var i = 0; i < n; i++) {
    var h = calcWorkHours_(times[i][0], times[i][1]);
    if (h === '') { unreadable++; continue; }        // 時刻が読めない行は触らない
    if (Number(cur[i][0]) === h) continue;           // 一致している行も触らない
    changed.push([i + 2, h]);
    diffs.push('  ' + (i + 2) + '行目 ' + formatDateCell_(dates[i][0]) + ' ' +
               names[i][0] + '：' + cur[i][0] + ' → ' + h);
  }

  var tail = unreadable > 0 ? '\n\n※ 開始・終了が読めない行が' + unreadable + '件あり、そのままにしました。' : '';
  if (diffs.length === 0) {
    ui.alert('勤務時間はすべて開始・終了と一致しています。' + tail);
    return;
  }

  var shown = diffs.slice(0, 20).join('\n');
  var more = diffs.length > 20 ? '\n  …ほか' + (diffs.length - 20) + '件' : '';
  var res = ui.alert('勤務時間の再計算',
    diffs.length + '件の勤務時間を書き換えます。よろしいですか？\n\n' + shown + more + tail,
    ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;

  changed.forEach(function(c) { sheet.getRange(c[0], COL.HOURS).setValue(c[1]); });
  ui.alert('勤務時間を再計算しました（' + diffs.length + '件を更新）。\n' +
           '管理用スプレッドシートで「割り振り → 業務完了報告より最新データを読み込み」も実行してください。');
}
