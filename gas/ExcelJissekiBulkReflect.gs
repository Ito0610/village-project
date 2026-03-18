/**
 * Excel実績 請求・目的コード・派遣人数 一括反映
 *
 * 【使い方】
 * 1. 利用者様名が入っているセル範囲を複数選択する
 * 2. メニュー「【一括反映】目的コード・派遣人数・請求」を実行
 *
 * 【ルール】
 * ◆請求（9列右のセル）
 * - 選択セルの背景が #fce5cd → スルー（何も記入しない）
 * - 選択セルの背景が #ff9900 または #ffff00 → 「居宅」を入力
 * - 上記以外で「Excel実績自動化フラグ」I列（3行目以降）に同じ利用者様名があれば「大田区」を入力
 * - 「大田区」を入力した行のみ、以下を実行
 *
 * ◆目的コード（4列右が目的地・参照、7列右に出力）
 * - 目的地が「Excel実績自動化フラグ」E列（3行目以降）のいずれかと部分一致 → 「F」
 * - それ以外 → 「A」
 *
 * ◆派遣人数（8列右のセル）
 * - 利用者様名が「Excel実績自動化フラグ」A列（3行目以降）のいずれかと一致 かつ
 *   1列左のヘルパー名が2人以上 → 「2」
 * - 上記でヘルパーが1人だけ、または利用者様名がリストにない → 「1」
 *
 * ◆スキップ条件
 * - 7列右（目的コード）・8列右（派遣人数）・9列右（請求）のいずれかに既に値が入っている場合は、そのセルは処理しない
 */

var BULK_REFLECT_CONFIG = {
  FLAG_SHEET_NAME: 'Excel実績自動化フラグ',
  SKIP_COLOR_HEX: '#fce5cd',
  KYOTAKU_COLOR_HEXES: ['#ff9900', '#ffff00'],
  I_COLUMN: 9,
  E_COLUMN: 5,
  A_COLUMN: 1,
  FIRST_ROW: 3
};

/**
 * メニューに「【一括反映】目的コード・派遣人数・請求」を追加する
 * 既存の onOpen() から addBulkReflectMenu() を呼び出してください。
 */
function addBulkReflectMenu() {
  SpreadsheetApp.getUi()
    .createMenu('【一括反映】目的コード・派遣人数・請求')
    .addItem('一括反映を実行', 'applyBulkReflect')
    .addToUi();
}

/**
 * 選択範囲（利用者様名のセル）に対して請求・目的コード・派遣人数を一括反映する
 */
function applyBulkReflect() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = ss.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert('利用者様名が入っているセル範囲を選択してから実行してください。');
    return;
  }

  var startColumn = range.getColumn();
  if (startColumn < 2) {
    SpreadsheetApp.getUi().alert('選択範囲はB列以降で指定してください。\n（1列左のヘルパー名を参照するため）');
    return;
  }

  var flagSheet = ss.getSheetByName(BULK_REFLECT_CONFIG.FLAG_SHEET_NAME);
  if (!flagSheet) {
    SpreadsheetApp.getUi().alert('シート「' + BULK_REFLECT_CONFIG.FLAG_SHEET_NAME + '」が見つかりません。');
    return;
  }

  var skipHex = normalizeBulkHex(BULK_REFLECT_CONFIG.SKIP_COLOR_HEX);
  var kyotakuHexes = BULK_REFLECT_CONFIG.KYOTAKU_COLOR_HEXES.map(function(h) { return normalizeBulkHex(h); });

  var iColumnList = getColumnList(ss, BULK_REFLECT_CONFIG.I_COLUMN);
  var eColumnList = getColumnList(ss, BULK_REFLECT_CONFIG.E_COLUMN);
  var aColumnList = getColumnList(ss, BULK_REFLECT_CONFIG.A_COLUMN);

  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var startRow = range.getRow();

  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      var row = startRow + r;
      var col = range.getColumn() + c;
      var nameCell = sheet.getRange(row, col);

      var billingCol = col + 9;
      var purposeCol = col + 7;
      var dispatchCol = col + 8;
      if (cellHasValue(sheet.getRange(row, purposeCol)) ||
          cellHasValue(sheet.getRange(row, dispatchCol)) ||
          cellHasValue(sheet.getRange(row, billingCol))) {
        continue;
      }

      var bg = normalizeBulkHex(nameCell.getBackground());

      if (bg === skipHex) {
        continue;
      }

      var userDisplay = nameCell.getDisplayValue();
      var userStr = (userDisplay != null ? String(userDisplay).trim() : '');

      var destCol = col + 4;
      var helperCol = col - 1;

      if (kyotakuHexes.indexOf(bg) !== -1) {
        sheet.getRange(row, billingCol).setValue('居宅');
        continue;
      }

      if (isInListExact(userStr, iColumnList)) {
        sheet.getRange(row, billingCol).setValue('大田区');

        var destCell = sheet.getRange(row, destCol);
        var destStr = (destCell.getDisplayValue() != null ? String(destCell.getDisplayValue()).trim() : '');
        var purposeValue = isInListPartial(destStr, eColumnList) ? 'F' : 'A';
        sheet.getRange(row, purposeCol).setValue(purposeValue);

        var isInDispatchList = isInListExact(userStr, aColumnList);
        var helperCell = sheet.getRange(row, helperCol);
        var helperStr = helperCell.getDisplayValue();
        var personCount = countHelperPersonsBulk(helperStr);
        var dispatchValue = (isInDispatchList && personCount >= 2) ? '2' : '1';
        sheet.getRange(row, dispatchCol).setValue(dispatchValue);
      }
    }
  }
}

/**
 * セルに値が入っているか（表示値が空白でないか）を判定する
 */
function cellHasValue(cell) {
  var v = cell.getDisplayValue();
  if (v == null) return false;
  return String(v).trim() !== '';
}

/**
 * 指定列（3行目以降）の表示値をリストで取得（空白は除く）
 */
function getColumnList(ss, columnIndex) {
  var flagSheet = ss.getSheetByName(BULK_REFLECT_CONFIG.FLAG_SHEET_NAME);
  if (!flagSheet) return [];

  var lastRow = flagSheet.getLastRow();
  if (lastRow < BULK_REFLECT_CONFIG.FIRST_ROW) return [];

  var values = flagSheet.getRange(BULK_REFLECT_CONFIG.FIRST_ROW, columnIndex, lastRow, columnIndex).getDisplayValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    var v = (values[i][0] != null ? String(values[i][0]).trim() : '');
    if (v !== '') list.push(v);
  }
  return list;
}

function isInListExact(str, list) {
  if (!str) return false;
  for (var i = 0; i < list.length; i++) {
    if (str === list[i]) return true;
  }
  return false;
}

function isInListPartial(str, list) {
  if (!str) return false;
  for (var i = 0; i < list.length; i++) {
    if (list[i] !== '' && str.indexOf(list[i]) !== -1) return true;
  }
  return false;
}

/**
 * ヘルパー名文字列を全角・半角スペースで区切った人数を返す
 */
function countHelperPersonsBulk(helperStr) {
  if (helperStr == null || (typeof helperStr !== 'string')) return 1;
  var s = String(helperStr).trim();
  if (s === '') return 1;
  var parts = s.split(/[\s\u3000]+/);
  var count = 0;
  for (var i = 0; i < parts.length; i++) {
    if (parts[i].trim() !== '') count++;
  }
  return count >= 1 ? count : 1;
}

function normalizeBulkHex(hex) {
  if (hex == null || hex === '') return '';
  var s = String(hex).trim().toLowerCase();
  if (s.indexOf('#') !== 0) s = '#' + s;
  return s;
}
