/**
 * 売上データ 月次サマリー＆ダッシュボード
 *
 * 「売上データ」シートを月ごとに集計し、「月次サマリー」シートに書き込んだうえで
 * 月次推移の棒グラフ（縦棒）を作成する。
 * setDailyTrigger を一度実行すると、毎朝9時台に runDashboard が自動実行される。
 */

// ===== 設定 =====
const SOURCE_SHEET_NAME = '売上データ';   // 集計元シート名
const SUMMARY_SHEET_NAME = '月次サマリー'; // 出力先シート名
const TRIGGER_HOUR = 9;                   // 自動実行する時刻（時）
const TRIGGER_TIMEZONE = 'Asia/Tokyo';    // トリガーのタイムゾーン

/**
 * メイン処理：売上データを月ごとに集計し、サマリーとグラフを更新する。
 */
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!source) {
    throw new Error(`「${SOURCE_SHEET_NAME}」シートが見つかりません。`);
  }

  // 出力先シートがなければ作成する
  const summary = ss.getSheetByName(SUMMARY_SHEET_NAME) || ss.insertSheet(SUMMARY_SHEET_NAME);

  // 全行を一括で読み込み、月ごとに集計する
  const rows = source.getDataRange().getValues();
  const monthly = aggregateByMonth_(rows, ss.getSpreadsheetTimeZone());

  // サマリーシートをクリアしてから書き直す
  writeSummary_(summary, monthly);

  // グラフを作り直す
  rebuildChart_(summary, monthly.length);

  console.log(`集計完了：${monthly.length}か月分を「${SUMMARY_SHEET_NAME}」に出力しました。`);
}

/**
 * 毎朝9時台に runDashboard を自動実行するトリガーを登録する。
 * 既存の runDashboard トリガーは削除してから登録するため、何度実行しても重複しない。
 */
function setDailyTrigger() {
  // 既存の同じトリガーを削除（重複登録の防止）
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'runDashboard')
    .forEach((t) => ScriptApp.deleteTrigger(t));

  // 毎日9時台に実行（GASの仕様上、9:00〜10:00の間のいずれかの時刻に実行される）
  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .inTimezone(TRIGGER_TIMEZONE)
    .create();

  console.log(`毎日${TRIGGER_HOUR}時台に runDashboard を実行するトリガーを登録しました。`);
}

/**
 * 売上データの行を月ごとに集計する。
 * 日付や金額として読み取れない行（見出し行・空行など）はスキップする。
 *
 * @param {Array<Array<*>>} rows シートの全行（A:日付, B:担当者, C:商品, D:金額）
 * @param {string} timeZone 日付を解釈するタイムゾーン
 * @return {Array<{year: number, month: number, total: number, count: number}>} 古い月から順の集計結果
 */
function aggregateByMonth_(rows, timeZone) {
  const map = {};

  rows.forEach((row) => {
    const ym = parseYearMonth_(row[0], timeZone);
    const amount = parseAmount_(row[3]);
    if (!ym || amount === null) return; // 見出し行・空行・不正データは除外

    const key = ym.year * 100 + ym.month; // 例：202601（並べ替え用のキー）
    if (!map[key]) {
      map[key] = { year: ym.year, month: ym.month, total: 0, count: 0 };
    }
    map[key].total += amount;
    map[key].count += 1;
  });

  // 古い月から順に並べる
  return Object.keys(map)
    .map(Number)
    .sort((a, b) => a - b)
    .map((key) => map[key]);
}

/**
 * セルの値から年・月を取り出す。日付型と「2026/01/05」形式の文字列に対応する。
 *
 * @param {*} value セルの値
 * @param {string} timeZone タイムゾーン
 * @return {{year: number, month: number}|null} 解釈できない場合は null
 */
function parseYearMonth_(value, timeZone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const ym = Utilities.formatDate(value, timeZone, 'yyyy/M').split('/');
    return { year: Number(ym[0]), month: Number(ym[1]) };
  }
  if (typeof value === 'string') {
    const m = value.trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-]\d{1,2}/);
    if (m) {
      const month = Number(m[2]);
      if (month >= 1 && month <= 12) return { year: Number(m[1]), month: month };
    }
  }
  return null;
}

/**
 * セルの値を金額（数値）に変換する。「1,000」のようなカンマ付き文字列にも対応する。
 *
 * @param {*} value セルの値
 * @return {number|null} 数値にできない場合は null
 */
function parseAmount_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : null;
  if (typeof value === 'string' && value.trim() !== '') {
    const n = Number(value.replace(/[,，円¥\s]/g, ''));
    return isFinite(n) ? n : null;
  }
  return null;
}

/**
 * 月次サマリーシートをクリアし、集計結果を書き込む。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 出力先シート
 * @param {Array<{year: number, month: number, total: number, count: number}>} monthly 集計結果
 */
function writeSummary_(sheet, monthly) {
  sheet.clear(); // 値と書式をクリア

  const values = [['月', '合計売上', '件数']];
  monthly.forEach((m) => {
    values.push([`${m.year}年${m.month}月`, m.total, m.count]);
  });

  // 「2026年1月」が日付に自動変換されないよう、A列を書式なしテキストにしておく
  sheet.getRange(1, 1, values.length, 1).setNumberFormat('@');
  sheet.getRange(1, 1, values.length, 3).setValues(values);

  // 見た目の調整
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f1f3f4');
  if (monthly.length > 0) {
    sheet.getRange(2, 2, monthly.length, 1).setNumberFormat('#,##0');
    sheet.getRange(2, 3, monthly.length, 1).setNumberFormat('#,##0');
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

/**
 * 月次サマリーシート上の既存グラフを削除し、合計売上の月次推移グラフを作り直す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 月次サマリーシート
 * @param {number} monthCount 集計した月の数
 */
function rebuildChart_(sheet, monthCount) {
  // sheet.clear() ではグラフは消えないため、個別に削除する
  sheet.getCharts().forEach((chart) => sheet.removeChart(chart));

  if (monthCount === 0) return; // データがなければグラフは作らない

  // A列（月）とB列（合計売上）を見出し行込みでグラフ化する
  const range = sheet.getRange(1, 1, monthCount + 1, 2);
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN) // 縦棒グラフ
    .addRange(range)
    .setNumHeaders(1)
    .setPosition(2, 5, 0, 0) // E2セルの位置に配置
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上（円）', format: '#,##0' })
    .setOption('width', 720)
    .setOption('height', 400)
    .build();

  sheet.insertChart(chart);
}
