/**
 * 掃除ルーレット — スプレッドシートバインド型 GAS Web アプリ
 *
 * 裏側のスプレッドシートを DB として扱い、
 *   ・「掃除場所」シート … 掃除場所の一覧 (A 列)
 *   ・「スタッフ」シート   … スタッフ名の一覧 (A 列)
 * を編集するだけで、ルーレットの中身を変更できる。
 */

// シート名（DB のテーブル名にあたる）
const SHEET_PLACES = '掃除場所';
const SHEET_STAFF = 'スタッフ';
const SHEET_HISTORY = '履歴';

// シートが存在しない初回起動時に投入する初期データ（デザインの既定値）
const DEFAULT_PLACES = ['技術部会議室', '多目的ホール１', '多目的ホール２', '貴賓応接間', '応接間', '外の休憩室', 'カフェエリア'];
const DEFAULT_STAFF = ['今泉', '笹生', '岩間', '奥', '日暮', '箱崎', '市田'];

/**
 * スプレッドシートを開いた時にカスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧹 掃除ルーレット')
    .addItem('マスタシートを初期化（無ければ作成）', 'setupSheets')
    .addToUi();
}

/**
 * Web アプリのエントリポイント。
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('掃除ルーレット 🧹')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 指定名のシートを取得し、無ければ初期データ付きで作成して返す。
 */
function getOrCreateSheet_(ss, name, header, defaults) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1).setValue(header).setFontWeight('bold');
  if (defaults && defaults.length) {
    const values = defaults.map(function (v) { return [v]; });
    sheet.getRange(2, 1, values.length, 1).setValues(values);
  }
  sheet.setColumnWidth(1, 220);
  return sheet;
}

/**
 * マスタシート（掃除場所・スタッフ）を用意する。
 * メニューからの初期化、および getMasterData からの遅延生成で利用。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet_(ss, SHEET_PLACES, '掃除場所', DEFAULT_PLACES);
  getOrCreateSheet_(ss, SHEET_STAFF, 'スタッフ', DEFAULT_STAFF);
  getOrCreateSheet_(ss, SHEET_HISTORY, '実行日時', []);

  // 履歴シートのヘッダを整える（初回のみ）
  const hist = ss.getSheetByName(SHEET_HISTORY);
  if (hist && hist.getRange(1, 2).getValue() === '') {
    hist.getRange(1, 1, 1, 3)
      .setValues([['実行日時', '掃除場所', '担当スタッフ']])
      .setFontWeight('bold');
  }
}

/**
 * A 列の値（ヘッダ行を除く）を空白を飛ばして配列で返す。
 */
function readColumn_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, 1).getValues();
  return values
    .map(function (row) { return String(row[0]).trim(); })
    .filter(function (v) { return v !== ''; });
}

/**
 * クライアントから呼ばれる。掃除場所とスタッフをスプレッドシートから読み出す。
 */
function getMasterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_PLACES) || !ss.getSheetByName(SHEET_STAFF)) {
    setupSheets();
  }
  return {
    places: readColumn_(ss.getSheetByName(SHEET_PLACES)),
    staff: readColumn_(ss.getSheetByName(SHEET_STAFF))
  };
}

/**
 * ルーレット結果を履歴シートに追記する。
 * assignments: [{ place: string, staff: string }, ...]
 */
function saveResult(assignments) {
  if (!assignments || !assignments.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hist = ss.getSheetByName(SHEET_HISTORY);
  if (!hist) {
    setupSheets();
    hist = ss.getSheetByName(SHEET_HISTORY);
  }
  const now = new Date();
  const rows = assignments.map(function (a) {
    return [now, a.place, a.staff];
  });
  hist.getRange(hist.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}
