/**
 * Excel実績 1/2 入力マクロ
 *
 * 【使い方】
 * 1. 値を入れたいセル範囲を選択する（例：H3:H100）
 * 2. メニュー「1/2入力」→「選択範囲に1/2を反映」を実行
 *
 * 【ルール】
 * - 各セルについて「6つ左」のセル（同じ行・左に6列）を参照する（利用者様名）
 * - 参照セル（6つ左）の背景が「水色」の場合のみ、次のように値を入れる
 *   - 参照セルの値が「Excel実績自動化フラグ」シート A列3行目以降の利用者様名と一致する場合：
 *     - 「7つ左」のセル（ヘルパー名）を確認。スペースで区切られた人数が2人以上 → 半角「2」
 *     - ヘルパーが1人だけのとき → 半角「1」
 *   - 利用者様名がリストに一致しない場合 → 半角「1」
 * - 6つ左のセルが水色でない場合は、何も書き込まない（既存の値はそのまま）
 *
 * 【水色の設定】
 * メニュー「1/2入力」→「選択セルの色コードを表示」で、実際に使っている水色のセルを
 * 選択して実行すると16進コードが表示されます。その値を下記 CONFIG.LIGHT_BLUE_HEX に設定してください。
 */

var CONFIG_12 = {
  // 水色とみなす背景色（16進）。シートで使っている水色に合わせて変更する
  LIGHT_BLUE_HEX: '#00FFFF',
  // 利用者様名が並んでいるシート名
  NAMES_SHEET_NAME: 'Excel実績自動化フラグ',
  // 利用者様名の列（A列 = 1）
  NAMES_COLUMN: 1,
  // 利用者様名の開始行（3行目から）
  NAMES_FIRST_ROW: 3
};

/**
 * スプレッドシートを開いたときに「1/2入力」メニューを追加する
 * - このスクリプトだけを使う場合：この関数名を onOpen に変更してください
 * - F/A入力と両方使う場合：ExcelJissekiFAInput.gs の onOpen に
 *   ui.createMenu('1/2入力')
 *     .addItem('選択範囲に1/2を反映', 'apply12ToSelection')
 *     .addItem('選択セルの色コードを表示', 'showSelectedCellColor12')
 *     .addToUi();
 *   を追加してください（README_ExcelJisseki12.md 参照）
 */
function onOpen12() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('1/2入力')
    .addItem('選択範囲に1/2を反映', 'apply12ToSelection')
    .addItem('選択セルの色コードを表示', 'showSelectedCellColor12')
    .addToUi();
}

/**
 * 選択範囲に対して、6つ左が水色なら 1 または 2 を入力する
 */
function apply12ToSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = ss.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert('セル範囲を選択してから実行してください。');
    return;
  }

  var startColumn = range.getColumn();
  if (startColumn < 8) {
    SpreadsheetApp.getUi().alert('選択範囲はH列より右で指定してください。\n（6つ左・7つ左のセルを参照するため、最低でもH列（8列目）以降が必要です）');
    return;
  }

  var namesList = getClientNamesList(ss);
  var lightBlueHex = normalizeHexColor12(CONFIG_12.LIGHT_BLUE_HEX);

  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var startRow = range.getRow();

  var toSet1 = [];
  var toSet2 = [];

  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      var targetRow = startRow + r;
      var targetCol = range.getColumn() + c;
      var refCol6 = targetCol - 6;  // 利用者様名
      var refCol7 = targetCol - 7;  // ヘルパー名

      var refCell6 = sheet.getRange(targetRow, refCol6);
      var bg = normalizeHexColor12(refCell6.getBackground());
      if (bg !== lightBlueHex) {
        continue; // 水色でない場合は何もしない
      }

      var refValue6 = refCell6.getDisplayValue();
      var refStr6 = (refValue6 != null ? String(refValue6).trim() : '');
      var isMatch = isInClientNamesList(refStr6, namesList);

      if (isMatch) {
        var refCell7 = sheet.getRange(targetRow, refCol7);
        var helperStr = refCell7.getDisplayValue();
        var personCount = countHelperPersons(helperStr);
        if (personCount >= 2) {
          toSet2.push(sheet.getRange(targetRow, targetCol));
        } else {
          toSet1.push(sheet.getRange(targetRow, targetCol));
        }
      } else {
        toSet1.push(sheet.getRange(targetRow, targetCol));
      }
    }
  }

  for (var i = 0; i < toSet1.length; i++) {
    toSet1[i].setValue('1');
  }
  for (var j = 0; j < toSet2.length; j++) {
    toSet2[j].setValue('2');
  }

  var msg = '処理しました。';
  if (toSet1.length > 0 || toSet2.length > 0) {
    msg += ' 1:' + toSet1.length + '件、 2:' + toSet2.length + '件';
  } else {
    msg += ' 水色の参照セルで該当したものはありませんでした。';
  }
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 利用者様名リストを取得（Excel実績自動化フラグ の A列3行目以降、空白は飛ばす）
 * @param {Spreadsheet} ss
 * @return {Array.<string>} トリム済み・空でない文字列の配列
 */
function getClientNamesList(ss) {
  var namesSheet = ss.getSheetByName(CONFIG_12.NAMES_SHEET_NAME);
  if (!namesSheet) {
    throw new Error('シート「' + CONFIG_12.NAMES_SHEET_NAME + '」が見つかりません。');
  }

  var lastRow = namesSheet.getLastRow();
  if (lastRow < CONFIG_12.NAMES_FIRST_ROW) {
    return [];
  }

  var col = CONFIG_12.NAMES_COLUMN;
  var values = namesSheet.getRange(CONFIG_12.NAMES_FIRST_ROW, col, lastRow, col).getDisplayValues();
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
 * 文字列が利用者様名リストに含まれるか（表示値で比較）
 */
function isInClientNamesList(str, namesList) {
  if (!str || str === '') return false;
  for (var i = 0; i < namesList.length; i++) {
    if (namesList[i] === str) return true;
  }
  return false;
}

/**
 * ヘルパー名セルの文字列から「人数」をカウントする
 * 全角スペース・半角スペースで区切られた名前の個数とする
 * @param {string} helperStr ヘルパー名（例：「市川　奥原(初)　荻原」）
 * @return {number} 人数（1以上）
 */
function countHelperPersons(helperStr) {
  if (helperStr == null || (typeof helperStr !== 'string')) return 1;
  var s = String(helperStr).trim();
  if (s === '') return 1;
  // 全角スペース(\u3000)と半角スペースで分割し、空でない要素の数
  var parts = s.split(/[\s\u3000]+/);
  var count = 0;
  for (var i = 0; i < parts.length; i++) {
    if (parts[i].trim() !== '') count++;
  }
  return count >= 1 ? count : 1;
}

/**
 * 背景色を #RRGGBB 形式に正規化（比較用）
 */
function normalizeHexColor12(hex) {
  if (hex == null || hex === '') return '';
  var s = String(hex).trim().toLowerCase();
  if (s.indexOf('#') !== 0) s = '#' + s;
  return s;
}

/**
 * 選択しているセルの背景色の16進コードを表示（水色設定の確認用）
 */
function showSelectedCellColor12() {
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!range || range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    SpreadsheetApp.getUi().alert('色を確認したいセルを1つだけ選択してから実行してください。');
    return;
  }
  var bg = range.getBackground();
  var hex = bg ? bg : '(なし)';
  SpreadsheetApp.getUi().alert('選択セルの背景色コード:\n' + hex + '\n\nこの値を CONFIG_12.LIGHT_BLUE_HEX に設定してください。');
}
