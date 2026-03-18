/**
 * Excel実績 請求フラグ入力
 *
 * 【使い方】
 * 1. 値を入れたいセル範囲を選択する（例：O3:O100）
 * 2. メニュー「Excel自動化」→「請求フラグ」を実行
 *
 * 【ルール】
 * 選択した各セルについて「7つ左」のセル（同じ行・左に7列）を参照する。
 * - 参照セルの背景が #ff9900（オレンジ） → 選択セルに「居宅」を入力
 * - 参照セルの背景が #00ffff（水色）かつ 参照セルの値が「Excel実績自動化フラグ」I列3行目以降のいずれかと完全一致 → 「大田区」を入力
 * - 上記以外 → 選択セルを空にする
 *
 * メニュー「Excel自動化」は ExcelJissekiFAInput.gs の addExcelAutomationMenu() で追加されます。
 * 本ファイルをプロジェクトに追加し、addExcelAutomationMenu() が実行されていれば「請求フラグ」が表示されます。
 */

var BILLING_CONFIG = {
  // オレンジ背景 → 「居宅」
  ORANGE_HEX: '#ff9900',
  // 水色背景かつI列3行目以降のいずれかと一致 → 「大田区」
  CYAN_HEX: '#00ffff',
  // 一致判定用シート・列（I列=9）、開始行（3行目以降）
  FLAG_SHEET_NAME: 'Excel実績自動化フラグ',
  MATCH_COLUMN: 9,
  MATCH_FIRST_ROW: 3
};

/**
 * 請求フラグ：選択セルに対して「7つ左」のセルを参照し、背景色に応じて「居宅」または「大田区」を入力する
 */
function applyBillingFlagToSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = ss.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert('セル範囲を選択してから実行してください。');
    return;
  }

  var startColumn = range.getColumn();
  if (startColumn < 8) {
    SpreadsheetApp.getUi().alert('選択範囲はH列より右で指定してください。\n（7つ左のセルを参照するため、最低でもH列以降が必要です）');
    return;
  }

  var flagSheet = ss.getSheetByName(BILLING_CONFIG.FLAG_SHEET_NAME);
  if (!flagSheet) {
    SpreadsheetApp.getUi().alert('シート「' + BILLING_CONFIG.FLAG_SHEET_NAME + '」が見つかりません。');
    return;
  }

  var matchList = getBillingMatchList(ss);

  var orangeHex = normalizeBillingHexColor(BILLING_CONFIG.ORANGE_HEX);
  var cyanHex = normalizeBillingHexColor(BILLING_CONFIG.CYAN_HEX);

  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var startRow = range.getRow();

  var countKyotaku = 0;
  var countOota = 0;
  var countOther = 0;

  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      var targetRow = startRow + r;
      var targetCol = range.getColumn() + c;
      var refCol = targetCol - 7;

      var refCell = sheet.getRange(targetRow, refCol);
      var targetCell = sheet.getRange(targetRow, targetCol);
      var bg = normalizeBillingHexColor(refCell.getBackground());

      var valueToSet = '';
      if (bg === orangeHex) {
        valueToSet = '居宅';
        countKyotaku++;
      } else if (bg === cyanHex) {
        var refDisplay = refCell.getDisplayValue();
        var refStr = (refDisplay != null ? String(refDisplay).trim() : '');
        if (isInBillingMatchList(refStr, matchList)) {
          valueToSet = '大田区';
          countOota++;
        } else {
          countOther++;
        }
      } else {
        countOther++;
      }

      targetCell.setValue(valueToSet);
    }
  }

  var msg = '請求フラグを反映しました。';
  msg += ' 居宅:' + countKyotaku + '件、 大田区:' + countOota + '件';
  if (countOther > 0) {
    msg += '、 対象外:' + countOther + '件';
  }
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 「Excel実績自動化フラグ」のI列3行目以降の表示値をリストで取得する（空白行は除く）
 * @param {Spreadsheet} ss
 * @return {Array.<string>}
 */
function getBillingMatchList(ss) {
  var flagSheet = ss.getSheetByName(BILLING_CONFIG.FLAG_SHEET_NAME);
  if (!flagSheet) return [];

  var lastRow = flagSheet.getLastRow();
  if (lastRow < BILLING_CONFIG.MATCH_FIRST_ROW) return [];

  var col = BILLING_CONFIG.MATCH_COLUMN;
  var firstRow = BILLING_CONFIG.MATCH_FIRST_ROW;
  var values = flagSheet.getRange(firstRow, col, lastRow, col).getDisplayValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    var v = (values[i][0] != null ? String(values[i][0]).trim() : '');
    if (v !== '') {
      list.push(v);
    }
  }
  return list;
}

/**
 * 文字列がリストのいずれかと完全一致するか
 */
function isInBillingMatchList(refStr, matchList) {
  for (var i = 0; i < matchList.length; i++) {
    if (refStr === matchList[i]) return true;
  }
  return false;
}

/**
 * 背景色を #RRGGBB 形式に正規化（比較用）
 */
function normalizeBillingHexColor(hex) {
  if (hex == null || hex === '') return '';
  var s = String(hex).trim().toLowerCase();
  if (s.indexOf('#') !== 0) s = '#' + s;
  return s;
}
