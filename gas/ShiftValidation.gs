/**
 * シフト表・シフト希望 照合チェック（Google Apps Script）
 *
 * 「シフトチェック」シートの C4〜C9 に記載された週ごとのシート名を参照し、
 * 各シートの D列「#073763」背景色2箇所でシフト表/シフト希望の行範囲を判定。
 * 1シートあたり7日分（+23列ずつ）のシフトをチェックする。
 *
 * チェック内容（3種類）:
 *   1. シフト表にいるがシフト希望にいない → 赤・太字
 *   2. シフト希望時間外で支援が入っている → 赤・太字
 *   3. シフト希望にいるがシフト表にいない → 赤・太字
 *
 * 起動方法:
 *   - 「シフトチェック」シートの D4〜D9 に「確認」と入力 → 対応する週をチェック
 *   - 「シフトチェック」シートの E4〜E9 に「黒」と入力 → 対応する週のマークをリセット
 *   - メニュー「シフト照合」から週ごと or 全週一括で実行
 */

// ========== 設定 ==========

/** シフト表/シフト希望の行範囲を区切る背景色 */
var MARKER_COLOR = '#073763';

/** マーカー色を判定する基準列（D列 = 4） */
var MARKER_COLUMN = 4;

/** 1日あたりの列オフセット */
var DAY_COL_OFFSET = 23;

/** 1シートあたりの日数 */
var DAYS_PER_SHEET = 7;

/** 各日のシフト列オフセット（1日目の先頭列からの相対位置、0-based） */
var COL_OFFSET_HELPER = 0;  // D列 = ヘルパー名
var COL_OFFSET_START = 2;   // F列 = 開始時刻
var COL_OFFSET_END = 3;     // G列 = 終了時刻

/** 1日目の先頭列（D列 = 4） */
var DAY1_START_COL = 4;

/** 「シフトチェック」シートの設定 */
var TRIGGER_SHEET_NAME = 'シフトチェック';

/** 週番号 → シート名セル、確認セル、リセットセルの対応（1-based 行・列） */
var WEEK_CONFIG = [
  { sheetNameCell: [4, 3], checkCell: [4, 4], resetCell: [4, 5] },   // 1週目: C4, D4, E4
  { sheetNameCell: [5, 3], checkCell: [5, 4], resetCell: [5, 5] },   // 2週目: C5, D5, E5
  { sheetNameCell: [6, 3], checkCell: [6, 4], resetCell: [6, 5] },   // 3週目: C6, D6, E6
  { sheetNameCell: [7, 3], checkCell: [7, 4], resetCell: [7, 5] },   // 4週目: C7, D7, E7
  { sheetNameCell: [8, 3], checkCell: [8, 4], resetCell: [8, 5] },   // 5週目: C8, D8, E8
  { sheetNameCell: [9, 3], checkCell: [9, 4], resetCell: [9, 5] }    // 6週目: C9, D9, E9
];

/** 週ラベル */
var WEEK_LABELS = ['1週目', '2週目', '3週目', '4週目', '5週目', '6週目'];

// ========== スタイル ==========

var ERROR_STYLE = { foreground: '#ff0000', bold: true };
var NORMAL_STYLE = { foreground: '#000000', bold: false };

// ========== メニュー ==========

function addShiftValidationMenu(ui) {
  if (!ui) ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('シフト確認');
  for (var i = 0; i < WEEK_LABELS.length; i++) {
    menu.addItem(WEEK_LABELS[i] + 'チェック', 'runWeekCheck_' + i);
  }
  menu.addSeparator();
  menu.addItem('全週チェック', 'runAllWeeksCheck');
  menu.addSeparator();
  for (var j = 0; j < WEEK_LABELS.length; j++) {
    menu.addItem(WEEK_LABELS[j] + 'リセット', 'runWeekReset_' + j);
  }
  menu.addSeparator();
  menu.addItem('全週リセット', 'runAllWeeksReset');
  menu.addItem('選択セルの色コードを表示', 'showSelectedCellColor');
  menu.addSeparator();
  menu.addItem('ヘルパーシフト転送', 'transferSubmissionsToWeeklyCalendar');
  menu.addToUi();
}

// ========== メニュー用エントリ ==========

function runWeekCheck_0() { runWeekValidation(0); }
function runWeekCheck_1() { runWeekValidation(1); }
function runWeekCheck_2() { runWeekValidation(2); }
function runWeekCheck_3() { runWeekValidation(3); }
function runWeekCheck_4() { runWeekValidation(4); }
function runWeekCheck_5() { runWeekValidation(5); }

function runWeekReset_0() { runWeekResetByIndex(0); }
function runWeekReset_1() { runWeekResetByIndex(1); }
function runWeekReset_2() { runWeekResetByIndex(2); }
function runWeekReset_3() { runWeekResetByIndex(3); }
function runWeekReset_4() { runWeekResetByIndex(4); }
function runWeekReset_5() { runWeekResetByIndex(5); }

function runAllWeeksCheck() {
  var results = [];
  for (var i = 0; i < WEEK_CONFIG.length; i++) {
    var sheetName = getWeekSheetName_(i);
    if (!sheetName) continue;
    var count = runWeekValidation(i, true);
    results.push(WEEK_LABELS[i] + '(' + sheetName + '): エラー ' + count + ' 件');
  }
  if (results.length === 0) {
    SpreadsheetApp.getUi().alert('チェック対象のシート名が1つも入力されていません。\nシフトチェックシートの C4〜C9 にシート名を入力してください。');
    return;
  }
  SpreadsheetApp.getUi().alert('全週チェック完了\n\n' + results.join('\n'));
}

function runAllWeeksReset() {
  var count = 0;
  for (var i = 0; i < WEEK_CONFIG.length; i++) {
    var sheetName = getWeekSheetName_(i);
    if (!sheetName) continue;
    runWeekResetByIndex(i, true);
    count++;
  }
  if (count === 0) {
    SpreadsheetApp.getUi().alert('リセット対象のシート名が1つも入力されていません。');
    return;
  }
  SpreadsheetApp.getUi().alert('全週（' + count + '週分）のリセットが完了しました。');
}

// ========== onEdit トリガー ==========

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== TRIGGER_SHEET_NAME) return;

  var row = e.range.getRow();
  var col = e.range.getColumn();
  var val = (e.value != null ? String(e.value) : '').trim();
  if (!val) return;

  for (var i = 0; i < WEEK_CONFIG.length; i++) {
    var cfg = WEEK_CONFIG[i];

    // D列に「確認」→ チェック実行
    if (row === cfg.checkCell[0] && col === cfg.checkCell[1]) {
      if (val === '確認') {
        try {
          var count = runWeekValidation(i, true);
          e.range.setValue('完了(エラー' + count + '件)');
        } catch (err) {
          e.range.setValue('エラー: ' + (err.message || err));
        }
      }
      return;
    }

    // E列に「黒」→ リセット実行
    if (row === cfg.resetCell[0] && col === cfg.resetCell[1]) {
      if (val === '黒') {
        try {
          runWeekResetByIndex(i, true);
          e.range.setValue('リセット完了');
        } catch (err) {
          e.range.setValue('エラー: ' + (err.message || err));
        }
      }
      return;
    }
  }
}

// ========== 週ごとのシート名取得 ==========

/**
 * 「シフトチェック」シートから指定週のシート名を取得
 */
function getWeekSheetName_(weekIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggerSheet = ss.getSheetByName(TRIGGER_SHEET_NAME);
  if (!triggerSheet) return '';
  var cell = WEEK_CONFIG[weekIndex].sheetNameCell;
  var v = triggerSheet.getRange(cell[0], cell[1]).getValue();
  return (v != null ? String(v) : '').trim();
}

// ========== 行範囲検出（#073763 ベース） ==========

/**
 * D列を走査し、MARKER_COLOR の背景色を持つ行番号を上から順に返す
 */
function findMarkerRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var range = sheet.getRange(1, MARKER_COLUMN, lastRow, 1);
  var backgrounds = range.getBackgrounds();
  var markerNorm = normalizeColor_(MARKER_COLOR);
  var rows = [];
  for (var i = 0; i < backgrounds.length; i++) {
    if (normalizeColor_(backgrounds[i][0]) === markerNorm) {
      rows.push(i + 1);
    }
  }
  return rows;
}

/**
 * マーカー行からシフト表・シフト希望の行範囲を算出。
 * #073763 が2箇所ある前提：1個目〜2個目の間 = シフト表、2個目の下 = シフト希望。
 */
function buildShiftBlocks_(sheet) {
  var markers = findMarkerRows_(sheet);
  if (markers.length < 2) return null;

  var lastRow = sheet.getLastRow();
  var shiftTableStart = markers[0] + 1;
  var shiftTableEnd = markers[1] - 1;
  var shiftHopeStart = markers[1] + 1;
  var shiftHopeEnd = lastRow;

  if (shiftTableStart > shiftTableEnd) shiftTableEnd = shiftTableStart - 1;
  if (shiftHopeStart > lastRow) shiftHopeEnd = shiftHopeStart - 1;

  return {
    shiftTableStart: shiftTableStart,
    shiftTableEnd: shiftTableEnd,
    shiftHopeStart: shiftHopeStart,
    shiftHopeEnd: shiftHopeEnd
  };
}

// ========== 列位置計算 ==========

/**
 * dayIndex (0=1日目, 1=2日目, ... 6=7日目) からヘルパー名・開始・終了の列番号を返す
 */
function getDayColumns_(dayIndex) {
  var base = DAY1_START_COL + dayIndex * DAY_COL_OFFSET;
  return {
    helper: base + COL_OFFSET_HELPER,
    start: base + COL_OFFSET_START,
    end: base + COL_OFFSET_END
  };
}

// ========== メイン照合ロジック ==========

/**
 * 指定週のチェックを実行
 * @param {number} weekIndex 0-5（1週目〜6週目）
 * @param {boolean} silent true のとき alert を出さない
 * @return {number} 検出エラー数
 */
function runWeekValidation(weekIndex, silent) {
  var sheetName = getWeekSheetName_(weekIndex);
  if (!sheetName) {
    var msg = WEEK_LABELS[weekIndex] + ': シート名が入力されていません（C' + WEEK_CONFIG[weekIndex].sheetNameCell[0] + '）';
    if (!silent) SpreadsheetApp.getUi().alert(msg);
    throw new Error(msg);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    var msg2 = 'シート「' + sheetName + '」が見つかりません。';
    if (!silent) SpreadsheetApp.getUi().alert(msg2);
    throw new Error(msg2);
  }

  var blocks = buildShiftBlocks_(sheet);
  if (!blocks) {
    var msg3 = 'シート「' + sheetName + '」で背景色 ' + MARKER_COLOR + ' のマーカー行が2つ見つかりません。';
    if (!silent) SpreadsheetApp.getUi().alert(msg3);
    throw new Error(msg3);
  }

  var totalErrors = 0;

  for (var day = 0; day < DAYS_PER_SHEET; day++) {
    var cols = getDayColumns_(day);
    var dayErrors = validateOneDay_(sheet, blocks, cols);
    totalErrors += dayErrors;
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      WEEK_LABELS[weekIndex] + '(' + sheetName + ') チェック完了。\n検出エラー: ' + totalErrors + ' 件'
    );
  }

  return totalErrors;
}

/**
 * 1日分のシフト照合を実行
 */
function validateOneDay_(sheet, blocks, cols) {
  var errorCount = 0;

  var tableData = readColumnData_(sheet, blocks.shiftTableStart, blocks.shiftTableEnd, cols);
  var hopeData = readColumnData_(sheet, blocks.shiftHopeStart, blocks.shiftHopeEnd, cols);

  var tableEntries = parseEntries_(tableData);
  var hopeByName = parseHopeByName_(hopeData);

  var tableHelperCol = cols.helper;
  var tableErrorRows = {};
  var hopeErrorNames = {};

  // 1. シフト表にいるがシフト希望にいない
  if (tableEntries.length > 0) {
    var tableHelpers = {};
    tableEntries.forEach(function(e) { tableHelpers[e.helper] = true; });
    Object.keys(tableHelpers).forEach(function(name) {
      if (!hopeByName[name] || hopeByName[name].length === 0) {
        markRowsByName_(sheet, blocks.shiftTableStart, blocks.shiftTableEnd, tableHelperCol, tableData, name, ERROR_STYLE);
        for (var i = 0; i < tableData.length; i++) {
          if (trimCell_(tableData[i].helper) === name) tableErrorRows[i] = true;
        }
        errorCount++;
      }
    });
  }

  // 2. シフト希望時間外で支援が入っている
  for (var i = 0; i < tableEntries.length; i++) {
    var entry = tableEntries[i];
    if (!entry.helper) continue;
    var hopes = hopeByName[entry.helper];
    if (!hopes || hopes.length === 0) continue;
    var within = hopes.some(function(h) {
      return isTimeWithin_(entry.start, entry.end, h.start, h.end);
    });
    if (!within) {
      tableErrorRows[entry.rowIndex] = true;
      markSingleRow_(sheet, blocks.shiftTableStart + entry.rowIndex, tableHelperCol, ERROR_STYLE);
      errorCount++;
    }
  }

  // 3. シフト希望にいるがシフト表にいない
  Object.keys(hopeByName).forEach(function(name) {
    if (!name) return;
    var inTable = tableEntries.some(function(e) { return e.helper === name; });
    if (!inTable) {
      hopeErrorNames[name] = true;
      markRowsByName_(sheet, blocks.shiftHopeStart, blocks.shiftHopeEnd, tableHelperCol, hopeData, name, ERROR_STYLE);
      errorCount++;
    }
  });

  // エラーに該当しない箇所を黒・通常に戻す
  resetNonErrorRows_(sheet, blocks.shiftTableStart, tableHelperCol, tableData, tableErrorRows, NORMAL_STYLE);
  resetNonErrorNames_(sheet, blocks.shiftHopeStart, tableHelperCol, hopeData, hopeErrorNames, NORMAL_STYLE);

  return errorCount;
}

// ========== データ読み取り ==========

/**
 * 指定行範囲からヘルパー名・開始・終了を読み取る
 */
function readColumnData_(sheet, startRow, endRow, cols) {
  if (startRow > endRow) return [];
  var numRows = endRow - startRow + 1;

  var helperRange = sheet.getRange(startRow, cols.helper, numRows, 1);
  var startRange = sheet.getRange(startRow, cols.start, numRows, 1);
  var endRange = sheet.getRange(startRow, cols.end, numRows, 1);

  var helpers = helperRange.getValues();
  var starts = startRange.getValues();
  var ends = endRange.getValues();

  var result = [];
  for (var i = 0; i < numRows; i++) {
    result.push({
      helper: helpers[i][0],
      start: starts[i][0],
      end: ends[i][0]
    });
  }
  return result;
}

/**
 * シフト表のデータをパース
 */
function parseEntries_(data) {
  var list = [];
  for (var i = 0; i < data.length; i++) {
    if (!isHelperNameCell_(data[i].helper)) continue;
    var name = trimCell_(data[i].helper);
    if (!name) continue;
    list.push({
      helper: name,
      start: parseTime_(data[i].start),
      end: parseTime_(data[i].end),
      rowIndex: i
    });
  }
  return list;
}

/**
 * シフト希望のデータをヘルパー名でグループ化
 */
function parseHopeByName_(data) {
  var byName = {};
  for (var i = 0; i < data.length; i++) {
    if (!isHelperNameCell_(data[i].helper)) continue;
    var name = trimCell_(data[i].helper);
    if (!name) continue;
    var start = parseTime_(data[i].start);
    if (start == null) start = 0;
    var end = parseTime_(data[i].end);
    if (!byName[name]) byName[name] = [];
    byName[name].push({ start: start, end: end });
  }
  return byName;
}

// ========== マーキング ==========

/**
 * 指定範囲内で特定ヘルパー名の行にスタイルを適用
 */
function markRowsByName_(sheet, startRow, endRow, helperCol, data, helperName, style) {
  for (var i = 0; i < data.length; i++) {
    if (trimCell_(data[i].helper) === helperName) {
      sheet.getRange(startRow + i, helperCol)
        .setFontColor(style.foreground)
        .setFontWeight(style.bold ? 'bold' : 'normal');
    }
  }
}

/**
 * 1行だけスタイルを適用
 */
function markSingleRow_(sheet, row, helperCol, style) {
  sheet.getRange(row, helperCol)
    .setFontColor(style.foreground)
    .setFontWeight(style.bold ? 'bold' : 'normal');
}

/**
 * シフト表: エラーでない行を通常に戻す
 */
function resetNonErrorRows_(sheet, startRow, helperCol, data, errorRowIndices, style) {
  for (var i = 0; i < data.length; i++) {
    if (errorRowIndices[i]) continue;
    sheet.getRange(startRow + i, helperCol)
      .setFontColor(style.foreground)
      .setFontWeight(style.bold ? 'bold' : 'normal');
  }
}

/**
 * シフト希望: エラーでない名前の行を通常に戻す
 */
function resetNonErrorNames_(sheet, startRow, helperCol, data, errorNames, style) {
  for (var i = 0; i < data.length; i++) {
    var name = trimCell_(data[i].helper);
    if (errorNames[name]) continue;
    sheet.getRange(startRow + i, helperCol)
      .setFontColor(style.foreground)
      .setFontWeight(style.bold ? 'bold' : 'normal');
  }
}

// ========== リセット ==========

/**
 * 指定週のヘルパー名を黒色・通常に戻す
 */
function runWeekResetByIndex(weekIndex, silent) {
  var sheetName = getWeekSheetName_(weekIndex);
  if (!sheetName) {
    var msg = WEEK_LABELS[weekIndex] + ': シート名が入力されていません。';
    if (!silent) SpreadsheetApp.getUi().alert(msg);
    throw new Error(msg);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    var msg2 = 'シート「' + sheetName + '」が見つかりません。';
    if (!silent) SpreadsheetApp.getUi().alert(msg2);
    throw new Error(msg2);
  }

  var blocks = buildShiftBlocks_(sheet);
  if (!blocks) {
    var msg3 = 'シート「' + sheetName + '」でマーカー行が見つかりません。';
    if (!silent) SpreadsheetApp.getUi().alert(msg3);
    throw new Error(msg3);
  }

  for (var day = 0; day < DAYS_PER_SHEET; day++) {
    var cols = getDayColumns_(day);
    var helperCol = cols.helper;

    if (blocks.shiftTableStart <= blocks.shiftTableEnd) {
      var numTableRows = blocks.shiftTableEnd - blocks.shiftTableStart + 1;
      sheet.getRange(blocks.shiftTableStart, helperCol, numTableRows, 1)
        .setFontColor(NORMAL_STYLE.foreground)
        .setFontWeight('normal');
    }

    if (blocks.shiftHopeStart <= blocks.shiftHopeEnd) {
      var numHopeRows = blocks.shiftHopeEnd - blocks.shiftHopeStart + 1;
      sheet.getRange(blocks.shiftHopeStart, helperCol, numHopeRows, 1)
        .setFontColor(NORMAL_STYLE.foreground)
        .setFontWeight('normal');
    }
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(WEEK_LABELS[weekIndex] + '(' + sheetName + ') のヘルパー名を黒色・通常に戻しました。');
  }
}

// ========== ユーティリティ ==========

function normalizeColor_(str) {
  if (!str) return '';
  var s = String(str).trim().toLowerCase();
  if (s.indexOf('#') !== 0) s = '#' + s;
  return s;
}

function trimCell_(v) {
  return (v != null && v !== '') ? String(v).trim() : '';
}

function isHelperNameCell_(v) {
  if (v == null || v === '') return false;
  if (typeof v === 'number' && (v > 1000 || (v >= 0 && v < 1))) return false;
  if (v instanceof Date) return false;
  var s = String(v).trim();
  if (!s) return false;
  if (/GMT|標準時|^\s*(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\w+\s+\d{4}/i.test(s)) return false;
  if (/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/.test(s)) return false;
  return true;
}

function parseTime_(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') {
    if (v < 1 && v >= 0) return v * 24 * 60;
    return v;
  }
  if (v instanceof Date) {
    return v.getHours() * 60 + v.getMinutes();
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  return null;
}

function toMinutes_(v) {
  if (v == null) return null;
  if (typeof v === 'number') {
    if (v >= 0 && v < 1) return v * 24 * 60;
    return v;
  }
  return v;
}

function isTimeWithin_(start, end, hopeStart, hopeEnd) {
  if (hopeStart == null) return false;
  var s = toMinutes_(start);
  var e = toMinutes_(end);
  var hs = toMinutes_(hopeStart);
  var he = toMinutes_(hopeEnd);
  if (e == null) e = s;
  if (he == null) he = 24 * 60;
  return s >= hs && e <= he;
}

function showSelectedCellColor() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selection = sheet.getActiveRange();
  if (!selection) {
    SpreadsheetApp.getUi().alert('セルを選択してから実行してください。');
    return;
  }
  var color = selection.getBackground();
  SpreadsheetApp.getUi().alert('選択セルの背景色: ' + (color || '(なし)'));
}
