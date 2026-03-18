/**
 * シフト表・シフト希望 照合チェック（Google Apps Script）
 * 【完成版①】 2025年2月時点の完成版として保存
 *
 * 対象範囲は「行は色」「列は固定ブロック」で指定し、3種類の間違いを検知して赤・太字でマークする。
 * 行の軸：対象色1〜5は同じ色を2行ずつ、対象色6は1行のみ使用。
 * 実行のたびにシートのマーカー列（デフォルトはD列）を走査し、その時点の色で範囲を判定するため、
 * 行の増減で位置が変わっても対応できる。
 */

// ========== 設定（ここを編集してください） ==========

/** 対象行を区切るマーカー行の背景色（シートで塗った色の16進コード） */
var CONFIG = {
  /** 対象色1 の背景色（例: 薄い黄）#ffff00 など */
  COLOR_1: '#7f6000',
  /** 対象色2 */
  COLOR_2: '#783f04',
  /** 対象色3 */
  COLOR_3: '#660000',
  /** 対象色4 */
  COLOR_4: '#5b0f00',
  /** 対象色5 */
  COLOR_5: '#4c1130',
  /** 対象色6 */
  COLOR_6: '#073763',
  /** マーカー色を判定する基準列（D列 = 4）。この列のセル背景で「対象色」を判定 */
  MARKER_COLUMN: 4
};

/** 列ブロック定義： [開始列番号(1-based), 終了列番号] の配列。D=4, E=5, F=6, G=7, H=8 など */
var COLUMN_BLOCKS = [
  [4, 8],   // D-H  日曜
  [10, 14], // J-N  月曜
  [16, 20], // P-T  火曜
  [22, 26], // V-Z  水曜
  [28, 32], // AB-AF 木曜
  [34, 38], // AH-AL 金曜
  [40, 44]  // AN-AR 土曜
];

/** メニュー表示名（曜日別チェック） */
var DAY_LABELS = [
  '日曜チェック',
  '月曜チェック',
  '火曜チェック',
  '水曜チェック',
  '木曜チェック',
  '金曜チェック',
  '土曜チェック'
];

/** シフト表・シフト希望の列オフセット（ブロック先頭から1-based） */
var COL_HELPER = 1;  // ヘルパー名
var COL_USER = 2;    // 利用者様名
var COL_START = 3;   // 開始時間
var COL_END = 4;     // 終了時間
var COL_NOTE = 5;    // 備考

// ========== スタイル ==========
/** エラー時：赤・太字 */
var ERROR_STYLE = {
  foreground: '#ff0000',
  bold: true
};
/** エラーに該当しない場合：黒・通常 */
var NORMAL_STYLE = {
  foreground: '#000000',
  bold: false
};

// ========== メイン実行 ==========

/**
 * メニューに「シフト照合」を追加（曜日別で列ブロックごとに実行し、タイムアウトを防ぐ）
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('シフト照合');
  for (var i = 0; i < DAY_LABELS.length; i++) {
    menu.addItem(DAY_LABELS[i], 'runShiftValidationForColumn_' + i);
  }
  menu.addSeparator()
    .addItem('ブロック・範囲の診断（日曜列）', 'runDiagnosticDay1')
    .addItem('選択セルの色コードを表示', 'showSelectedCellColor')
    .addToUi();
}

/** 日曜チェック用エントリ */
function runShiftValidationForColumn_0() { runShiftValidationForColumn(0); }
/** 月曜チェック用エントリ */
function runShiftValidationForColumn_1() { runShiftValidationForColumn(1); }
/** 火曜チェック用エントリ */
function runShiftValidationForColumn_2() { runShiftValidationForColumn(2); }
/** 水曜チェック用エントリ */
function runShiftValidationForColumn_3() { runShiftValidationForColumn(3); }
/** 木曜チェック用エントリ */
function runShiftValidationForColumn_4() { runShiftValidationForColumn(4); }
/** 金曜チェック用エントリ */
function runShiftValidationForColumn_5() { runShiftValidationForColumn(5); }
/** 土曜チェック用エントリ */
function runShiftValidationForColumn_6() { runShiftValidationForColumn(6); }

/**
 * 指定した列ブロック（曜日）だけ照合チェックを実行する。
 * メニューから曜日別に呼ばれ、実行時間を分散してタイムアウトを防ぐ。
 * @param {number} columnIndex 0=日曜(D-H), 1=月曜(J-N), 2=火曜(P-T), 3=水曜(V-Z), 4=木曜(AB-AF), 5=金曜(AH-AL), 6=土曜(AN-AR)
 */
function runShiftValidationForColumn(columnIndex) {
  if (columnIndex < 0 || columnIndex >= COLUMN_BLOCKS.length) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var markerRows = getMarkerRowsByColor(sheet);
  if (!markerRows || Object.keys(markerRows).length === 0) {
    SpreadsheetApp.getUi().alert('対象色で塗ったマーカー行が検出できません。CONFIG の色コードと、シートのマーカー列（D列）の色を確認してください。');
    return;
  }

  var blocks = buildRowBlocks(markerRows);
  if (blocks.length === 0) {
    SpreadsheetApp.getUi().alert('シフト表・シフト希望のブロックを特定できません。対象色1〜6が、色1・色1・色2・色2…の順で並んでいるか確認してください。');
    return;
  }

  var colBlock = COLUMN_BLOCKS[columnIndex];
  var dayLabel = DAY_LABELS[columnIndex];
  var errorCount = 0;

  var lastRow = sheet.getLastRow();

  for (var b = 0; b < blocks.length; b++) {
    var block = blocks[b];
    var hopeStart = block.shiftHopeStart;
    var hopeEnd = block.shiftHopeEnd;
    // 希望範囲が逆転（色2が希望1行目にある等）している場合は、希望終了をシート最終行まで延長する
    if (hopeEnd < hopeStart && hopeStart <= lastRow) {
      hopeEnd = lastRow;
    }

    var shiftTableRange = getBlockRange(sheet, block.shiftTableStart, block.shiftTableEnd, colBlock[0], colBlock[1]);
    var shiftHopeRange = getBlockRange(sheet, hopeStart, hopeEnd, colBlock[0], colBlock[1]);
    // シフト希望が無いブロックはスキップ。シフト表が空でもシフト希望のエラー3は検知するためブロックは処理する
    if (!shiftHopeRange) continue;

    var shiftTableData, tableEntries, tableHelperCells;
    if (shiftTableRange && shiftTableRange.getNumRows() > 0) {
      shiftTableData = shiftTableRange.getValues();
      tableEntries = parseShiftTable(shiftTableData);
      tableHelperCells = getHelperNameCellsInBlock(sheet, block.shiftTableStart, block.shiftTableEnd, colBlock[0]);
    } else {
      shiftTableData = [];
      tableEntries = [];
      tableHelperCells = null;
    }

    var shiftHopeData = shiftHopeRange.getValues();
    var hopeHelperCells = getHelperNameCellsInBlock(sheet, hopeStart, hopeEnd, colBlock[0]);
    var hopeByName = parseShiftHope(shiftHopeData);

    // エラー該当を記録（シフト表は行インデックス、シフト希望は名前で「該当しない」を黒・通常に戻すため）
    var tableErrorRowIndices = {};
    var hopeErrorHelpers = {};

    // 1. シフト表にいるがシフト希望にいない → シフト表のそのヘルパー名（全行）を赤・太字（シフト表がある場合のみ）
    if (tableHelperCells && shiftTableData.length > 0) {
      var tableHelpers = {};
      tableEntries.forEach(function (e) { tableHelpers[e.helper] = true; });
      Object.keys(tableHelpers).forEach(function (name) {
        if (!hopeByName[name] || hopeByName[name].length === 0) {
          markHelperCellsInTable(tableHelperCells, shiftTableData, name, ERROR_STYLE);
          for (var ti = 0; ti < shiftTableData.length; ti++) {
            if (trimCell(shiftTableData[ti][0]) === name) tableErrorRowIndices[ti] = true;
          }
          errorCount++;
        }
      });
    }

    // 2. シフト希望時間外で支援が入っている → 時間外の行のヘルパー名セルだけを赤・太字（シフト表がある場合のみ）
    if (tableHelperCells) {
      for (var i = 0; i < tableEntries.length; i++) {
        var e = tableEntries[i];
        if (!e.helper) continue;
        var hopes = hopeByName[e.helper];
        if (!hopes || hopes.length === 0) continue;
        var within = hopes.some(function (h) {
          return isTimeWithin(e.start, e.end, h.start, h.end);
        });
        if (!within) {
          var r = e.rowIndex;
          tableErrorRowIndices[r] = true;
          markTableHelperCellAtRow(tableHelperCells, r, ERROR_STYLE);
          errorCount++;
        }
      }
    }

    // 3. シフト希望にいるがシフト表にいない → シフト希望のヘルパー名を赤・太字（シフト表が空でも実行）
    var hopeHelpers = Object.keys(hopeByName);
    hopeHelpers.forEach(function (name) {
      if (!name) return;
      var inTable = tableEntries.some(function (e) { return e.helper === name; });
      if (!inTable) {
        hopeErrorHelpers[name] = true;
        markHelperCellsInHope(hopeHelperCells, shiftHopeData, name, ERROR_STYLE);
        errorCount++;
      }
    });

    // 3つのエラーに該当しない箇所は黒色・通常文字に戻す（空白セルも黒・通常）
    if (tableHelperCells) {
      resetTableHelperCellsToNormal(tableHelperCells, shiftTableData, tableErrorRowIndices, NORMAL_STYLE);
    }
    resetHelperCellsToNormal(hopeHelperCells, shiftHopeData, hopeErrorHelpers, NORMAL_STYLE);
  }

  SpreadsheetApp.getUi().alert(dayLabel + ' 完了。検出したエラー箇所を赤・太字でマークしました。（検出数: ' + errorCount + '）');
}

/**
 * マーカー列を走査し、対象色の行番号を色ごとに取得
 * @return {Object} { '#ffff00': [5, 10], '#ffcc99': [16, 21], ... }
 */
function getMarkerRowsByColor(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var col = CONFIG.MARKER_COLUMN;
  var range = sheet.getRange(1, col, lastRow, col);
  var backgrounds = range.getBackgrounds();
  var colors = [CONFIG.COLOR_1, CONFIG.COLOR_2, CONFIG.COLOR_3, CONFIG.COLOR_4, CONFIG.COLOR_5, CONFIG.COLOR_6];
  var normalized = {};
  colors.forEach(function (c) { normalized[normalizeColor(c)] = c; });

  var byColor = {};
  for (var i = 0; i < backgrounds.length; i++) {
    var bg = backgrounds[i][0];
    if (!bg) continue;
    var key = normalizeColor(bg);
    if (normalized[key] === undefined) continue;
    var hex = normalized[key];
    if (!byColor[hex]) byColor[hex] = [];
    byColor[hex].push(i + 1); // 1-based row
  }

  colors.forEach(function (c) {
    if (byColor[c]) byColor[c].sort(function (a, b) { return a - b; });
  });
  return byColor;
}

/**
 * 色文字列を正規化（小文字・#あり）
 */
function normalizeColor(str) {
  if (!str) return '';
  var s = String(str).trim().toLowerCase();
  if (s.indexOf('#') !== 0) s = '#' + s;
  return s;
}

/**
 * マーカー列（D列）を走査し、背景色が付いているセルの「色コード → 行番号の配列」を返す。
 * 診断で「実際に読んだ色」を表示するために使用。白・無色は除外する。
 */
function getActualBackgroundsInMarkerColumn(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var col = CONFIG.MARKER_COLUMN;
  var range = sheet.getRange(1, col, lastRow, col);
  var backgrounds = range.getBackgrounds();
  var byKey = {};
  for (var i = 0; i < backgrounds.length; i++) {
    var bg = backgrounds[i][0];
    var key = normalizeColor(bg || '');
    if (!key || key === '#ffffff' || key === '#fff') continue;
    if (!byKey[key]) byKey[key] = [];
    byKey[key].push(i + 1);
  }
  var result = [];
  Object.keys(byKey).forEach(function (key) {
    var hex = key.indexOf('#') === 0 ? key : '#' + key;
    result.push({ hex: hex, rows: byKey[key].sort(function (a, b) { return a - b; }) });
  });
  result.sort(function (a, b) { return (a.rows[0] || 0) - (b.rows[0] || 0); });
  return result;
}

/**
 * マーカー行の情報から「シフト表」「シフト希望」の行範囲を算出。
 * 対象色1・対象色1の間＝シフト表、対象色1・対象色2の間＝シフト希望 … の繰り返し。
 * 次の色の先頭行が「希望の開始」より上にある場合（色2が希望1行目付近にある等）は、
 * シフト表を「次の色の2行前」で打ち切り、希望開始を次の色の先頭行にする。
 */
function buildRowBlocks(markerRows) {
  var blocks = [];
  var colors = [CONFIG.COLOR_1, CONFIG.COLOR_2, CONFIG.COLOR_3, CONFIG.COLOR_4, CONFIG.COLOR_5, CONFIG.COLOR_6];
  for (var i = 0; i < colors.length - 1; i++) {
    var c1 = colors[i];
    var c2 = colors[i + 1];
    var rows1 = markerRows[c1];
    var rows2 = markerRows[c2];
    if (!rows1 || rows1.length < 2 || !rows2 || rows2.length < 1) continue;

    var tableStart = rows1[0] + 1;
    var tableEnd = rows1[1] - 1;
    var hopeStart = rows1[1] + 1;
    var hopeEnd = rows2[0] - 1;

    // 次の色の先頭が「希望開始」より上 → 希望が実際には次の色の行から始まるとみなす
    if (rows2[0] < rows1[1]) {
      tableEnd = Math.min(tableEnd, rows2[0] - 2);   // シフト表は「次の色の2行前」まで
      hopeStart = rows2[0];                           // 希望開始＝次の色の先頭行
      hopeEnd = rows2.length >= 2 ? rows2[1] - 1 : rows2[0] - 1;
    }

    if (tableStart <= tableEnd || hopeStart <= hopeEnd) {
      blocks.push({
        shiftTableStart: tableStart,
        shiftTableEnd: tableEnd,
        shiftHopeStart: hopeStart,
        shiftHopeEnd: hopeEnd
      });
    }
  }
  return blocks;
}

/**
 * 指定した行・列範囲の Range を取得する。
 * GAS の getRange(行, 列, 行数, 列数) に合わせて、終了行・終了列から行数・列数を計算する。
 */
function getBlockRange(sheet, startRow, endRow, startCol, endCol) {
  if (startRow > endRow) return null;
  var numRows = endRow - startRow + 1;
  var numCols = endCol - startCol + 1;
  return sheet.getRange(startRow, startCol, numRows, numCols);
}

/**
 * シフト表の有効行をパース（ヘルパー名・開始・終了）。
 * ヘルパー名が空の行はスキップするため、rowIndex で範囲内の実際の行番号（0-based）を保持する。
 */
function parseShiftTable(data) {
  var list = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!isHelperNameCell(row[COL_HELPER - 1])) continue;
    var helper = trimCell(row[COL_HELPER - 1]);
    if (!helper) continue;
    var start = parseTime(row[COL_START - 1]);
    var end = parseTime(row[COL_END - 1]);
    list.push({ helper: helper, start: start, end: end, rowIndex: i });
  }
  return list;
}

/**
 * シフト希望をパース：ヘルパー名 → [{ start, end }, ...]
 * 開始時間が空欄の場合は 0:00 とみなす（分換算で 0）。
 */
function parseShiftHope(data) {
  var byName = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!isHelperNameCell(row[COL_HELPER - 1])) continue;
    var helper = trimCell(row[COL_HELPER - 1]);
    if (!helper) continue;
    var start = parseTime(row[COL_START - 1]);  // シフト希望は3列目が開始
    if (start == null) start = 0;               // 空欄の場合は 0:00 とみなす
    var end = parseTime(row[COL_END - 1]);     // 4列目が終了
    if (!byName[helper]) byName[helper] = [];
    byName[helper].push({ start: start, end: end });
  }
  return byName;
}

function trimCell(v) {
  return (v != null && v !== '') ? String(v).trim() : '';
}

/**
 * セル値を「ヘルパー名として扱うか」判定。日付・数値・日付風文字列は false。
 */
function isHelperNameCell(v) {
  if (v == null || v === '') return false;
  if (typeof v === 'number' && (v > 1000 || (v >= 0 && v < 1))) return false; // 数値（日付シリアル等）
  if (v instanceof Date) return false;
  var s = String(v).trim();
  if (!s) return false;
  if (/GMT|標準時|^\s*(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\w+\s+\d{4}/i.test(s)) return false; // 日付文字列
  if (/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/.test(s)) return false; // 日付形式
  return true;
}

/**
 * セル値を「分」で表す数値に変換（比較用）。日付オブジェクト・数値・"9:00" 形式に対応
 */
function parseTime(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') {
    if (v < 1 && v >= 0) return v * 24 * 60; // シートの時刻シリアル
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

/**
 * 支援時間 [start, end] が希望時間 [hopeStart, hopeEnd] のいずれかに含まれるか
 * 終了未記入の場合は開始のみ一致でも可とする
 */
function isTimeWithin(start, end, hopeStart, hopeEnd) {
  if (hopeStart == null) return false;
  var s = toMinutes(start);
  var e = toMinutes(end);
  var hs = toMinutes(hopeStart);
  var he = toMinutes(hopeEnd);
  if (e == null) e = s; // 終了未記入は開始と同じとみなす
  if (he == null) he = 24 * 60;
  return s >= hs && e <= he;
}

function toMinutes(v) {
  if (v == null) return null;
  if (typeof v === 'number') {
    if (v >= 0 && v < 1) return v * 24 * 60;
    return v;
  }
  return v;
}

/**
 * シフト表内で「ヘルパー名」セルだけ取得（行インデックスと一致させるため）
 */
function getHelperNameCellsInBlock(sheet, startRow, endRow, blockStartCol) {
  var col = blockStartCol + COL_HELPER - 1;
  var range = sheet.getRange(startRow, col, endRow, col);
  return range;
}

/**
 * シフト表で指定ヘルパー名のセルを赤・太字にする（その名前の行すべて）
 */
function markHelperCellsInTable(helperCellRange, tableData, helperName, style) {
  var values = helperCellRange.getValues();
  for (var i = 0; i < values.length; i++) {
    if (trimCell(values[i][0]) === helperName) {
      helperCellRange.getCell(i + 1, 1).setFontColor(style.foreground).setFontWeight(style.bold ? 'bold' : 'normal');
    }
  }
}

/**
 * シフト表の指定行（0-based）のヘルパー名セルだけにスタイルを適用する（エラー2：時間外の行のみ）
 */
function markTableHelperCellAtRow(helperCellRange, rowIndex0Based, style) {
  helperCellRange.getCell(rowIndex0Based + 1, 1)
    .setFontColor(style.foreground)
    .setFontWeight(style.bold ? 'bold' : 'normal');
}

/**
 * シフト希望で指定ヘルパー名のセルを赤・太字にする
 */
function markHelperCellsInHope(helperCellRange, hopeData, helperName, style) {
  var values = helperCellRange.getValues();
  for (var i = 0; i < values.length; i++) {
    if (trimCell(values[i][0]) === helperName) {
      helperCellRange.getCell(i + 1, 1).setFontColor(style.foreground).setFontWeight(style.bold ? 'bold' : 'normal');
    }
  }
}

/**
 * シフト表：エラーに該当しない行のヘルパー名セルを黒色・通常文字に戻す（行インデックスで判定）。
 * 空白セルも黒字・通常にそろえる。
 */
function resetTableHelperCellsToNormal(helperCellRange, data, tableErrorRowIndices, style) {
  for (var i = 0; i < data.length; i++) {
    if (tableErrorRowIndices[i]) continue; // エラー該当行はそのまま
    helperCellRange.getCell(i + 1, 1).setFontColor(style.foreground).setFontWeight(style.bold ? 'bold' : 'normal');
  }
}

/**
 * シフト希望：エラーに該当しないヘルパー名のセルを黒色・通常文字に戻す（名前で判定）。
 * 空白セルも黒字・通常にそろえる。
 */
function resetHelperCellsToNormal(helperCellRange, data, errorHelperSet, style) {
  for (var i = 0; i < data.length; i++) {
    var name = trimCell(data[i][0]);
    if (errorHelperSet[name]) continue; // エラー該当はそのまま
    helperCellRange.getCell(i + 1, 1).setFontColor(style.foreground).setFontWeight(style.bold ? 'bold' : 'normal');
  }
}

/**
 * 日曜列（D-H）でブロック・範囲の診断を行い、「照合診断」シートに結果を書き出す。
 * エラー3が検知されない原因の切り分けに使用する。
 */
function runDiagnosticDay1() {
  runDiagnosticForColumn(0);
}

function runDiagnosticForColumn(columnIndex) {
  var ss, sheet, diagSheet, dayLabel;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.getActiveSheet(); // データのシート（診断実行前に開いていたシート）
    dayLabel = DAY_LABELS[columnIndex];

    // 最初に「照合診断」シートを取得または作成し、すぐ表示する
    diagSheet = ss.getSheetByName('照合診断');
    if (!diagSheet) {
      diagSheet = ss.insertSheet('照合診断');
    }
    diagSheet.clear();
    diagSheet.getRange(1, 1).setValue('診断実行中...');
    diagSheet.activate(); // このシートを前面に出す

    var markerRows = getMarkerRowsByColor(sheet);
    if (!markerRows || Object.keys(markerRows).length === 0) {
      diagSheet.getRange(1, 1).setValue('【診断結果】マーカー行が検出できませんでした。');
      diagSheet.getRange(2, 1).setValue('D列（マーカー列）の背景色が CONFIG の色コードと一致しているか確認してください。');
      var actualColors = getActualBackgroundsInMarkerColumn(sheet);
      var outRow = 3;
      diagSheet.getRange(outRow, 1).setValue('■ CONFIGで指定している色コード（一致する行が無いと検出されません）:');
      outRow++;
      var configColors = [CONFIG.COLOR_1, CONFIG.COLOR_2, CONFIG.COLOR_3, CONFIG.COLOR_4, CONFIG.COLOR_5, CONFIG.COLOR_6];
      configColors.forEach(function (c, idx) {
        diagSheet.getRange(outRow, 1).setValue('  対象色' + (idx + 1) + ': ' + c);
        outRow++;
      });
      diagSheet.getRange(outRow, 1).setValue('■ D列で実際に読んだ背景色（この値に合わせて CONFIG を書き換えてください）:');
      outRow++;
      if (actualColors.length === 0) {
        diagSheet.getRange(outRow, 1).setValue('  （色付きのセルが1件もありません。D列の該当行に塗りつぶしをしましたか？）');
        outRow++;
      } else {
        actualColors.forEach(function (item) {
          diagSheet.getRange(outRow, 1).setValue('  ' + item.hex + ' → 行: ' + item.rows.join(', '));
          outRow++;
        });
      }
      diagSheet.getRange(outRow, 1).setValue('※メニュー「選択セルの色コードを表示」でマーカーセルを選択して実際のコードを確認し、ShiftValidation.gs の CONFIG に貼り付けてください。');
      diagSheet.autoResizeColumns(1, 1);
      SpreadsheetApp.getUi().alert(dayLabel + ' の診断を実行しました。\n「照合診断」シートに結果を書き出しました。\nマーカー行が検出できていません。\n\n照合診断シートの「D列で実際に読んだ背景色」を CONFIG に合わせるか、そのコードを CONFIG に設定してください。');
      return;
    }

    var blocks = buildRowBlocks(markerRows);
    if (blocks.length === 0) {
      diagSheet.getRange(1, 1).setValue('【診断結果】ブロックが0件です。');
      diagSheet.getRange(2, 1).setValue('対象色1〜6が「色1・色1・色2・色2…」の順で、各色が正しい行にあるか確認してください。');
      diagSheet.getRange(3, 1).setValue('検出したマーカー行（色ごとの行番号）:');
      var r = 4;
      var colors = [CONFIG.COLOR_1, CONFIG.COLOR_2, CONFIG.COLOR_3, CONFIG.COLOR_4, CONFIG.COLOR_5, CONFIG.COLOR_6];
      colors.forEach(function (c) {
        if (markerRows[c] && markerRows[c].length > 0) {
          diagSheet.getRange(r, 1).setValue('色: ' + c + ' → 行: ' + markerRows[c].join(', '));
          r++;
        }
      });
      diagSheet.autoResizeColumns(1, 1);
      SpreadsheetApp.getUi().alert(dayLabel + ' の診断を実行しました。\n「照合診断」シートに結果を書き出しました。\nブロックが0件のため、マーカー行の情報のみ表示しています。');
      return;
    }

    var colBlock = COLUMN_BLOCKS[columnIndex];
    var lastRow = sheet.getLastRow();

    diagSheet.getRange(1, 1, 1, 6).setValues([['ブロック', 'シフト表範囲', 'シフト希望範囲', '表の人数', '希望の人数', '希望のみ(エラー3)']]);
    diagSheet.getRange(1, 1, 1, 6).setFontWeight('bold');

    var row = 2;
    for (var b = 0; b < blocks.length; b++) {
      var block = blocks[b];
      var hopeStart = block.shiftHopeStart;
      var hopeEnd = block.shiftHopeEnd;
      if (hopeEnd < hopeStart && hopeStart <= lastRow) hopeEnd = lastRow;

      var shiftTableRange = getBlockRange(sheet, block.shiftTableStart, block.shiftTableEnd, colBlock[0], colBlock[1]);
      var shiftHopeRange = getBlockRange(sheet, hopeStart, hopeEnd, colBlock[0], colBlock[1]);

      var tableRangeStr = block.shiftTableStart + '〜' + block.shiftTableEnd;
      var hopeRangeStr = hopeStart + '〜' + hopeEnd;
      var tableCount = 0;
      var hopeNames = [];
      var error3Names = [];

      if (shiftHopeRange) {
        var shiftHopeData = shiftHopeRange.getValues();
        var hopeByName = parseShiftHope(shiftHopeData);
        hopeNames = Object.keys(hopeByName);

        var tableEntries = [];
        if (shiftTableRange && shiftTableRange.getNumRows() > 0) {
          var shiftTableData = shiftTableRange.getValues();
          tableEntries = parseShiftTable(shiftTableData);
          tableCount = tableEntries.length;
        }

        hopeNames.forEach(function (name) {
          if (!name) return;
          var inTable = tableEntries.some(function (e) { return e.helper === name; });
          if (!inTable) error3Names.push(name);
        });
      }

      diagSheet.getRange(row, 1, 1, 6).setValues([[
        b + 1,
        tableRangeStr,
        hopeRangeStr,
        tableCount,
        hopeNames.length,
        error3Names.join(', ') || '(なし)'
      ]]);
      row++;
    }

    diagSheet.autoResizeColumns(1, 6);
    SpreadsheetApp.getUi().alert(dayLabel + ' の診断を完了しました。\n「照合診断」シートにブロックごとの範囲とエラー3でマークされる名前を書き出しました。\n\nD109〜D112 が含まれるブロックの「シフト希望範囲」に 109〜112 が入っているか、「希望のみ(エラー3)」に 伊藤信一・稲山・荻原・佳織 が出ているかを確認してください。');
  } catch (e) {
    try {
      if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
      diagSheet = ss.getSheetByName('照合診断');
      if (!diagSheet) diagSheet = ss.insertSheet('照合診断');
      diagSheet.clear();
      diagSheet.getRange(1, 1).setValue('【診断でエラーが発生しました】');
      diagSheet.getRange(2, 1).setValue(String(e.message || e));
      diagSheet.getRange(3, 1).setValue('スタック: ' + (e.stack || ''));
      diagSheet.activate();
    } catch (e2) {}
    SpreadsheetApp.getUi().alert('診断中にエラーが発生しました。\n「照合診断」シートを確認するか、以下の内容を共有してください。\n\n' + (e.message || e));
  }
}

/**
 * 選択セルの背景色コードをログ＆トーストで表示（色設定の確認用）
 */
function showSelectedCellColor() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selection = sheet.getActiveRange();
  if (!selection) {
    SpreadsheetApp.getUi().alert('セルを選択してから実行してください。');
    return;
  }
  var color = selection.getBackground();
  var msg = '選択セルの背景色: ' + (color || '(なし)');
  SpreadsheetApp.getUi().alert(msg);
  Logger.log(msg);
}
