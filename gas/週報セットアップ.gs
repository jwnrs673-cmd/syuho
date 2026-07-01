/**
 * 週報スプレッドシート セットアップスクリプト（設計v1準拠）
 *
 * 実行するとアクティブなスプレッドシートに以下を構築します:
 *   - M_大分類 / M_担当者 / M_案件 / T_タスク / ダッシュボード の各シート
 *   - 各シートのヘッダー・書式（太字・固定行・列幅）
 *   - プルダウン（入力規則）… 大分類名・担当者名・案件名・区分・ステータス
 *   - M_案件.累計作業時間 の自動集計数式（T_タスクの実績を SUMIFS）
 *   - ダッシュボード … 大分類別の実績時間集計＋円グラフ＋ドリルダウン（案件別）
 *
 * 使い方:
 *   1. 対象スプレッドシートを開く →「拡張機能」→「Apps Script」
 *   2. このファイルの内容を貼り付けて保存
 *   3. 関数 setupWeeklyReport を実行（初回は承認が必要）
 *   4. 以後はメニュー「週報ツール」→「セットアップ実行」からも実行可
 *
 * 特徴:
 *   - 冪等（何度実行してもOK）。既存シートのデータは消さず、ヘッダー・入力規則・数式のみ再適用。
 *   - マスタ（大分類・担当者）は空のときだけサンプル行を投入。
 */

// ===== 設定：列定義（ヘッダー） ==========================================
var HEADERS = {
  'M_大分類': ['大分類ID', '大分類名', '表示順', '色'],
  'M_担当者': ['担当者ID', '担当者名', '所属', '表示順', '在籍'],
  'M_案件': ['案件ID', '案件名', '大分類名', '担当者名', 'ステータス', '進捗率',
    '取り組み内容', 'コメント／課題', '開始日', '期限', '完了日', '累計作業時間'],
  'T_タスク': ['タスクID', '案件名', '区分', '担当者名', '大分類名', '日付',
    '作業内容', '時間(h)', 'ステータス', 'コメント／課題']
};

// マスタが空のときだけ投入するサンプル
var SEED_大分類 = [
  ['K01', '営業・商談', 1, '#4285F4'],
  ['K02', '電話対応', 2, '#EA4335'],
  ['K03', '資料作成', 3, '#FBBC04'],
  ['K04', '社内会議', 4, '#34A853'],
  ['K05', 'メール対応', 5, '#FF6D01'],
  ['K06', 'その他', 9, '#9E9E9E']
];
var SEED_担当者 = [
  ['U01', '担当者A', '営業1課', 1, true],
  ['U02', '担当者B', '営業1課', 2, true]
];

var STATUS_案件 = ['未着手', '進行中', '完了', '保留'];
var STATUS_タスク = ['未着手', '進行中', '完了'];
var 区分LIST = ['予定', '実績'];

var HEADER_BG = '#1155cc';
var HEADER_FG = '#ffffff';

// ===== メニュー ==========================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('週報ツール')
    .addItem('セットアップ実行', 'setupWeeklyReport')
    .addToUi();
}

// ===== メイン ============================================================
function setupWeeklyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var shDaibunrui = ensureSheet_(ss, 'M_大分類', HEADERS['M_大分類']);
  var shTanto = ensureSheet_(ss, 'M_担当者', HEADERS['M_担当者']);
  var shAnken = ensureSheet_(ss, 'M_案件', HEADERS['M_案件']);
  var shTask = ensureSheet_(ss, 'T_タスク', HEADERS['T_タスク']);

  // マスタのサンプル投入（空のときだけ）
  seedIfEmpty_(shDaibunrui, SEED_大分類);
  seedIfEmpty_(shTanto, SEED_担当者);

  // 名前付き範囲（プルダウン参照用。1000行まで拡張余地を持たせる）
  setNamedRange_(ss, 'LIST_大分類', shDaibunrui, 2, 2, 1000, 1); // B2:B1001
  setNamedRange_(ss, 'LIST_担当者', shTanto, 2, 2, 1000, 1);     // B2:B1001
  setNamedRange_(ss, 'LIST_案件', shAnken, 2, 2, 1000, 1);       // B2:B1001

  applyMasterFormatting_(shDaibunrui, shTanto);
  applyAnkenSheet_(shAnken, shDaibunrui, shTanto);
  applyTaskSheet_(shTask, shAnken, shTanto, shDaibunrui);
  buildDashboard_(ss, shDaibunrui);

  SpreadsheetApp.getUi().alert('週報スプレッドシートのセットアップが完了しました。');
}

// ===== シート生成 =========================================================
function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  // ヘッダー書き込み＆装飾（データ行は触らない）
  var range = sh.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), headers.length); // ensure width exists
  return sh;
}

function seedIfEmpty_(sh, rows) {
  var lastRow = sh.getLastRow();
  if (lastRow > 1) return; // 既にデータあり
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// ===== 名前付き範囲 =======================================================
function setNamedRange_(ss, name, sh, row, col, numRows, numCols) {
  var range = sh.getRange(row, col, numRows, numCols);
  // 既存の同名を消して再設定
  var existing = ss.getNamedRanges();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getName() === name) existing[i].remove();
  }
  ss.setNamedRange(name, range);
}

// ===== 書式・入力規則 =====================================================
function applyMasterFormatting_(shDaibunrui, shTanto) {
  shDaibunrui.setColumnWidth(2, 160);
  shTanto.setColumnWidth(2, 140);
  // 在籍列をチェックボックス化
  var last = Math.max(shTanto.getLastRow(), 2);
  shTanto.getRange(2, 5, 1000, 1).insertCheckboxes();
}

function applyAnkenSheet_(sh, shDaibunrui, shTanto) {
  var ss = sh.getParent();
  var maxRows = 1000;

  // 大分類名（C列）: LIST_大分類
  setValidationFromNamed_(sh, 3, maxRows, ss.getRangeByName('LIST_大分類'));
  // 担当者名（D列）: LIST_担当者
  setValidationFromNamed_(sh, 4, maxRows, ss.getRangeByName('LIST_担当者'));
  // ステータス（E列）
  setValidationFromList_(sh, 5, maxRows, STATUS_案件);
  // 進捗率（F列）: 0-100 の数値。%表示
  sh.getRange(2, 6, maxRows, 1).setNumberFormat('0"%"');
  // 日付列（I,J,K）
  sh.getRange(2, 9, maxRows, 3).setNumberFormat('yyyy-mm-dd');
  // 累計作業時間（L列）: T_タスクの実績を案件名で集計
  var formulaRows = maxRows;
  var formulas = [];
  for (var r = 2; r <= formulaRows + 1; r++) {
    formulas.push(['=IF($A' + r + '="","",SUMIFS(\'T_タスク\'!$H:$H,\'T_タスク\'!$B:$B,$B' + r +
      ',\'T_タスク\'!$C:$C,"実績"))']);
  }
  sh.getRange(2, 12, formulaRows, 1).setFormulas(formulas);
  sh.getRange(2, 12, formulaRows, 1).setNumberFormat('0.0');

  // 列幅
  sh.setColumnWidth(2, 200); // 案件名
  sh.setColumnWidth(7, 240); // 取り組み内容
  sh.setColumnWidth(8, 240); // コメント
}

function applyTaskSheet_(sh, shAnken, shTanto, shDaibunrui) {
  var ss = sh.getParent();
  var maxRows = 2000;

  // 案件名（B列）: LIST_案件
  setValidationFromNamed_(sh, 2, maxRows, ss.getRangeByName('LIST_案件'));
  // 区分（C列）
  setValidationFromList_(sh, 3, maxRows, 区分LIST);
  // 担当者名（D列）
  setValidationFromNamed_(sh, 4, maxRows, ss.getRangeByName('LIST_担当者'));
  // 大分類名（E列）
  setValidationFromNamed_(sh, 5, maxRows, ss.getRangeByName('LIST_大分類'));
  // 日付（F列）
  sh.getRange(2, 6, maxRows, 1).setNumberFormat('yyyy-mm-dd');
  // 時間(h)（H列）
  sh.getRange(2, 8, maxRows, 1).setNumberFormat('0.0');
  // ステータス（I列）
  setValidationFromList_(sh, 9, maxRows, STATUS_タスク);

  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(7, 260);
  sh.setColumnWidth(10, 220);
}

function setValidationFromNamed_(sh, col, numRows, namedRange) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(namedRange, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, col, numRows, 1).setDataValidation(rule);
}

function setValidationFromList_(sh, col, numRows, list) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, col, numRows, 1).setDataValidation(rule);
}

// ===== ダッシュボード =====================================================
function buildDashboard_(ss, shDaibunrui) {
  var name = 'ダッシュボード';
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.clearFormats();
  removeChartsOfSheet_(sh);

  // 期間フィルタ
  sh.getRange('A1').setValue('期間開始').setFontWeight('bold');
  sh.getRange('A2').setValue('期間終了').setFontWeight('bold');
  sh.getRange('B1:B2').setNumberFormat('yyyy-mm-dd');
  sh.getRange('D1').setValue('選択大分類（ドリルダウン）').setFontWeight('bold');
  // 選択大分類プルダウン
  var ruleK = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRangeByName('LIST_大分類'), true)
    .setAllowInvalid(true).build();
  sh.getRange('E1').setDataValidation(ruleK);

  // 大分類別 実績時間 集計（円グラフの元）
  sh.getRange('A4').setValue('■ 大分類別 実績時間（円グラフ）').setFontWeight('bold');
  sh.getRange('A5').setValue('大分類').setFontWeight('bold').setBackground('#eeeeee');
  sh.getRange('B5').setValue('実績時間(h)').setFontWeight('bold').setBackground('#eeeeee');

  // 大分類名を master から参照して並べる（1000行対応）
  var catCount = Math.max(shDaibunrui.getLastRow() - 1, 1);
  var startRow = 6;
  var catFormulas = [];
  var sumFormulas = [];
  for (var i = 0; i < catCount; i++) {
    var srcRow = i + 2; // M_大分類の行
    catFormulas.push(['=IF(\'M_大分類\'!$B' + srcRow + '="","",\'M_大分類\'!$B' + srcRow + ')']);
    var dr = startRow + i;
    sumFormulas.push(['=IF($A' + dr + '="","",SUMIFS(\'T_タスク\'!$H:$H,' +
      '\'T_タスク\'!$E:$E,$A' + dr + ',' +
      '\'T_タスク\'!$C:$C,"実績",' +
      '\'T_タスク\'!$F:$F,">="&IF($B$1="",DATE(1900,1,1),$B$1),' +
      '\'T_タスク\'!$F:$F,"<="&IF($B$2="",DATE(2999,12,31),$B$2)))']);
  }
  sh.getRange(startRow, 1, catCount, 1).setFormulas(catFormulas);
  sh.getRange(startRow, 2, catCount, 1).setFormulas(sumFormulas);
  sh.getRange(startRow, 2, catCount, 1).setNumberFormat('0.0');

  // 円グラフ作成
  var pieRange = sh.getRange(5, 1, catCount + 1, 2); // ヘッダー含む A5:B(...)
  var pie = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(pieRange)
    .setPosition(4, 4, 0, 0) // 行4,列4(D)付近
    .setOption('title', '大分類別 実績時間')
    .setOption('pieHole', 0)
    .setOption('width', 460)
    .setOption('height', 300)
    .build();
  sh.insertChart(pie);

  // ドリルダウン：選択大分類の 案件別 実績時間
  var ddRow = startRow + catCount + 3;
  sh.getRange(ddRow - 1, 1).setValue('■ ドリルダウン：選択大分類の 案件別 実績時間').setFontWeight('bold');
  // QUERY で案件別集計（E1で選んだ大分類、期間はB1/B2）
  var query = '=IFERROR(QUERY(\'T_タスク\'!B:H,' +
    '"select B, sum(H) where C=\'実績\' ' +
    'and E=\'"&$E$1&"\' group by B order by sum(H) desc label B \'案件\', sum(H) \'実績時間(h)\'",1),' +
    '"該当データなし")';
  sh.getRange(ddRow, 1).setFormula(query);

  sh.setColumnWidth(1, 200);
  sh.autoResizeColumn(2);
  ss.setActiveSheet(sh);
}

function removeChartsOfSheet_(sh) {
  var charts = sh.getCharts();
  for (var i = 0; i < charts.length; i++) sh.removeChart(charts[i]);
}
