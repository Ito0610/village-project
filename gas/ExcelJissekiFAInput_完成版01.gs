/**
 * Excel実績  F/A 入力マクロ
 * 完成版01
 *
 * 【使い方】
 * 1. 値を入れたいセル範囲を選択する（例：G3:G100）
 * 2. メニュー「Excel自動化」→「目的コード反映」を実行
 *
 * 【ルール】
 * - 各セルについて「2つ左」のセル（同じ行・左に2列）を参照する
 * - 参照セルの背景が「水色」の場合：
 *   - 参照セルの値が「Excel実績自動化フラグ」シート E列3行目以降のいずれかと部分一致 → 半角「F」
 *   - Fの対象ではない全ての水色のセル（空・リストにない等） → 「A」
 * - 参照セルが水色でない場合は、選択セルを空にする
 *
 * 【水色の設定】
 * メニュー「Excel自動化」→「選択セルの色コードを表示」で、実際に使っている水色のセルを
 * 選択して実行すると16進コードが表示されます。その値を下記 CONFIG.LIGHT_BLUE_HEX に設定してください。
 */

var CONFIG = {
  // 水色とみなす背景色（16進）。シートで使っている水色に合わせて変更する
  LIGHT_BLUE_HEX: '#00FFFF',
  // 「ある言葉」が並んでいるシート名
  WORDS_SHEET_NAME: 'Excel実績自動化フラグ',
  // 「ある言葉」の列（E列 = 5）
  WORDS_COLUMN: 5,
  // 「ある言葉」の開始行（3行目から）
  WORDS_FIRST_ROW: 3
};

/**
 * 「Excel自動化」メニューをメニューバーに追加する
 * スプレッドシートを開いたときにメニューを出すには、プロジェクトのどこかで
 * onOpen() を定義し、その中で addExcelAutomationMenu(); を呼んでください。
 *
 * 例（Excel自動化だけ使う場合）：
 *   function onOpen() {
 *     addExcelAutomationMenu();
 *   }
 *
 * 例（1/2入力など他のメニューも使う場合）：
 *   function onOpen() {
 *     addExcelAutomationMenu();
 *     var ui = SpreadsheetApp.getUi();
 *     ui.createMenu('1/2入力')...
 *   }
 *
 * ※ メニューを手で addItem で並べず、必ず addExcelAutomationMenu() を呼ぶこと。
 *    そうしないと「請求フラグ」など追加した項目がメニューに反映されません。
 */
function addExcelAutomationMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Excel自動化')
    .addItem('目的コード反映', 'applyFAToSelection')
    .addItem('2人付反映', 'apply12ToSelection')
    .addItem('請求フラグ', 'applyBillingFlagToSelection')
    .addItem('【一括反映】目的コード・派遣人数・請求', 'applyBulkReflect')
    .addItem('選択セルの色コードを表示', 'showSelectedCellColor')
    .addToUi();
}

/**
 * 選択範囲に対して、2つ左のセルが水色なら F または A を入力する
 */
function applyFAToSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = ss.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert('セル範囲を選択してから実行してください。');
    return;
  }

  var startColumn = range.getColumn();
  if (startColumn < 3) {
    SpreadsheetApp.getUi().alert('選択範囲はC列より右で指定してください。\n（2つ左のセルを参照するため、最低でもC列以降が必要です）');
    return;
  }

  var wordsList = getWordsList(ss);
  var lightBlueHex = normalizeHexColor(CONFIG.LIGHT_BLUE_HEX);

  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var startRow = range.getRow();

  var toSetF = [];
  var toSetA = [];
  var toSetEmpty = [];

  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      var targetRow = startRow + r;
      var targetCol = range.getColumn() + c;
      var refCol = targetCol - 2;

      var refCell = sheet.getRange(targetRow, refCol);
      var targetCell = sheet.getRange(targetRow, targetCol);
      var bg = normalizeHexColor(refCell.getBackground());

      if (bg !== lightBlueHex) {
        toSetEmpty.push(targetCell);
        continue;
      }

      var refValue = refCell.getValue();
      var refDisplay = refCell.getDisplayValue();

      if (isInWordsList(refValue, refDisplay, wordsList)) {
        toSetF.push(targetCell);
      } else {
        // Fの対象ではない全ての水色のセルにAを入力（空・リストにない等を含む）
        toSetA.push(targetCell);
      }
    }
  }

  for (var i = 0; i < toSetF.length; i++) {
    toSetF[i].setValue('F');
  }
  for (var j = 0; j < toSetA.length; j++) {
    toSetA[j].setValue('A');
  }
  for (var k = 0; k < toSetEmpty.length; k++) {
    toSetEmpty[k].setValue('');
  }

  var msg = '処理しました。';
  msg += ' F:' + toSetF.length + '件、 A:' + toSetA.length + '件、 空:' + toSetEmpty.length + '件';
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 「ある言葉」リストを取得（Excel実績自動化フラグ の E列3行目以降）
 * @param {Spreadsheet} ss
 * @return {Array.<string>} トリム済み・空でない文字列の配列
 */
function getWordsList(ss) {
  var wordsSheet = ss.getSheetByName(CONFIG.WORDS_SHEET_NAME);
  if (!wordsSheet) {
    throw new Error('シート「' + CONFIG.WORDS_SHEET_NAME + '」が見つかりません。');
  }

  var lastRow = wordsSheet.getLastRow();
  if (lastRow < CONFIG.WORDS_FIRST_ROW) {
    return [];
  }

  var col = CONFIG.WORDS_COLUMN;
  var values = wordsSheet.getRange(CONFIG.WORDS_FIRST_ROW, col, lastRow, col).getDisplayValues();
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
 * 値が「ある言葉」リストのいずれかと部分一致するか（参照セルの値にリストの語が含まれていれば true）
 */
function isInWordsList(value, displayValue, wordsList) {
  var s = displayValue != null && String(displayValue).trim() !== ''
    ? String(displayValue).trim()
    : (value != null ? String(value).trim() : '');
  if (s === '') return false;
  for (var i = 0; i < wordsList.length; i++) {
    if (wordsList[i] !== '' && s.indexOf(wordsList[i]) !== -1) return true;
  }
  return false;
}

/**
 * 背景色を #RRGGBB 形式に正規化（比較用）
 */
function normalizeHexColor(hex) {
  if (hex == null || hex === '') return '';
  var s = String(hex).trim().toLowerCase();
  if (s.indexOf('#') !== 0) s = '#' + s;
  return s;
}

/**
 * 選択しているセルの背景色の16進コードを表示（水色設定の確認用）
 */
function showSelectedCellColor() {
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!range || range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    SpreadsheetApp.getUi().alert('色を確認したいセルを1つだけ選択してから実行してください。');
    return;
  }
  var bg = range.getBackground();
  var hex = bg ? bg : '(なし)';
  SpreadsheetApp.getUi().alert('選択セルの背景色コード:\n' + hex + '\n\nこの値を CONFIG.LIGHT_BLUE_HEX に設定してください。');
}
