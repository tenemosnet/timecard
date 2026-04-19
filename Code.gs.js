/**
 * ====================================
 *  勤怠管理システム — Google Apps Script  ver3.0
 * ====================================
 *
 *  セットアップ手順:
 *  1. Googleスプレッドシートを新規作成
 *  2. メニュー「拡張機能 → Apps Script」を開く
 *  3. このコードを Code.gs に貼り付けて保存
 *  4. 同じプロジェクト内に「Index.html」を作成し、HTMLコードを貼り付け
 *  5. スプレッドシートをリロード → メニュー「勤怠管理 → 初期設定」を実行
 *  6. 「デプロイ → 新しいデプロイ → Webアプリ」で公開
 *     - 実行するユーザー: 自分
 *     - アクセスできるユーザー: 全員
 *  7. 発行されたURLをタブレットのブラウザで開く
 *
 *  給与締め日: 毎月15日（16日〜翌15日が1ヶ月の勤務期間）
 *  例: 「4月」= 3/16〜4/15
 */

// =====================================================================
//  設定（ここを書き換えてカスタマイズ）
// =====================================================================
const SPREADSHEET_ID = '16ssw5Jfu3y8CepLMVMk_rnI4zBJBFUg8pnlfPYvFEPY';
const STAFF_NAMES = ['飯島祥子', '中井川雪子', '安達あけみ', '佐藤涼子']; // ※ 初期設定時のみ使用。以降はシート名で管理
const LOG_SHEET_NAME = '打刻ログ';
const NOTES_LOG_SHEET_NAME = '備考ログ';
const SETTINGS_SHEET_NAME = 'スタッフ設定';
const HOLIDAYS_SHEET_NAME = '祝日';
const PDF_ROOT_FOLDER_NAME = '勤怠管理PDF';
const SYSTEM_SHEET_NAMES = [LOG_SHEET_NAME, NOTES_LOG_SHEET_NAME, SETTINGS_SHEET_NAME, HOLIDAYS_SHEET_NAME];
const HOLIDAY_CALENDAR_ID = 'ja.japanese#holiday@group.v.calendar.google.com';
const DEFAULT_CONTRACTED_HOURS = 8;  // デフォルト定時（時間）
const BREAK_HOURS = 1;               // 休憩時間（固定1時間）
const NIGHT_START = 22;              // 深夜開始 22:00
const NIGHT_END = 5;                 // 深夜終了 5:00
const CUTOFF_DAY = 15;               // 給与締め日

// =====================================================================
//  メニュー
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('勤怠管理')
    .addItem('初期設定（初回のみ）', 'initialize')
    .addSeparator()
    .addItem('月を移行', 'promptNewMonth')
    .addItem('PDF出力（月次）', 'exportPdf')
    .addItem('備考を復元', 'restoreNotes')
    .addSeparator()
    .addItem('曜日の数式を修復', 'fixWeekdayFormulas')
    .addItem('罫線を一括適用', 'applyBordersToAll')
    .addItem('数式を一括更新', 'updateAllFormulas')
    .addItem('数式を個別更新（1名）', 'updateSingleStaffFormulas')
    .addItem('祝日を更新', 'updateHolidays')
    .addItem('祝日トリガーを設定', 'setupHolidayTrigger')
    .addItem('スタッフ名を変更', 'showStaffHelp')
    .addToUi();
}

// =====================================================================
//  Webアプリ: 打刻画面を表示
// =====================================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('タイムカード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================================
//  初期設定
// =====================================================================
function initialize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '初期設定',
    '打刻ログシートとスタッフ別シート（' + STAFF_NAMES.length + '名分）を作成します。\n続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // タイムゾーンを日本に設定
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  // 打刻ログシート作成
  createLogSheet_(ss);

  // 備考ログシート作成
  createNotesLogSheet_(ss);

  // スタッフ設定シート作成
  createSettingsSheet_(ss);

  // 祝日シート作成 + 当年の祝日取得 + 自動トリガー設定
  createHolidaysSheet_(ss);
  const now = new Date();
  updateHolidaysSheet_(ss, now.getFullYear());
  setupHolidayTrigger();

  // スタッフ別シート作成（当月）
  STAFF_NAMES.forEach(name => {
    createStaffSheet_(ss, name, now.getFullYear(), now.getMonth() + 1);
    setStaffSetting_(ss, name, DEFAULT_CONTRACTED_HOURS);
  });

  // デフォルトの空シートを削除
  ['シート1', 'Sheet1'].forEach(defaultName => {
    const s = ss.getSheetByName(defaultName);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (e) { /* 無視 */ }
    }
  });

  ui.alert(
    '初期設定が完了しました',
    STAFF_NAMES.length + '名分のシートを作成しました。\n\n' +
    '次のステップ:\n' +
    '「デプロイ → 新しいデプロイ → Webアプリ」で公開してください。\n' +
    '  • 実行するユーザー: 自分\n' +
    '  • アクセスできるユーザー: 全員',
    ui.ButtonSet.OK
  );
}

// =====================================================================
//  打刻ログシート作成
// =====================================================================
function createLogSheet_(ss) {
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME, 0);
  } else {
    return sheet;
  }

  sheet.getRange('A1:D1').setValues([['記録日時', '氏名', '種別', '日付']]);
  sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#e2e8f0');
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 110);
  sheet.setFrozenRows(1);

  return sheet;
}

// =====================================================================
//  備考ログシート作成
// =====================================================================
function createNotesLogSheet_(ss) {
  let sheet = ss.getSheetByName(NOTES_LOG_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(NOTES_LOG_SHEET_NAME, 1);
  sheet.getRange('A1:E1').setValues([['年', '月', 'スタッフ名', '日', '備考内容']]);
  sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e2e8f0');
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 40);
  sheet.setColumnWidth(5, 300);
  sheet.setFrozenRows(1);
  return sheet;
}

// =====================================================================
//  スタッフ設定シート作成
// =====================================================================
function createSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SETTINGS_SHEET_NAME, 2);
  sheet.getRange('A1:C1').setValues([['スタッフ名', '定時(時間)', '表示順']]);
  sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#e2e8f0');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 80);
  sheet.setFrozenRows(1);
  return sheet;
}

// =====================================================================
//  祝日シート作成
// =====================================================================
function createHolidaysSheet_(ss) {
  let sheet = ss.getSheetByName(HOLIDAYS_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(HOLIDAYS_SHEET_NAME);
  sheet.getRange('A1:B1').setValues([['日付', '祝日名']]);
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#e2e8f0');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd');
  sheet.setFrozenRows(1);
  return sheet;
}

// =====================================================================
//  祝日取得（Google Calendar API）
// =====================================================================
// Google Calendarに含まれるが国民の祝日ではないイベント
const NON_OFFICIAL_HOLIDAYS = [
  'ひな祭り', '雛祭り', '七五三', '七夕', 'クリスマス', 'バレンタインデー',
  'ホワイトデー', 'ハロウィン', '大晦日', '節分',
  '母の日', '父の日', 'クリスマス イブ'
];

function fetchHolidays_(year) {
  const cal = CalendarApp.getCalendarById(HOLIDAY_CALENDAR_ID);
  if (!cal) throw new Error('日本の祝日カレンダーにアクセスできません。');

  const start = new Date(year, 0, 1);   // 1/1
  const end = new Date(year, 11, 31, 23, 59, 59); // 12/31
  const events = cal.getEvents(start, end);

  return events
    .filter(ev => !NON_OFFICIAL_HOLIDAYS.includes(ev.getTitle()))
    .map(ev => ({
      date: ev.getAllDayStartDate(),
      name: ev.getTitle()
    }));
}

// =====================================================================
//  祝日シート更新（指定年の祝日を書き込み）
// =====================================================================
function updateHolidaysSheet_(ss, year) {
  let sheet = ss.getSheetByName(HOLIDAYS_SHEET_NAME);
  if (!sheet) sheet = createHolidaysSheet_(ss);

  // 既存データを読み込み、指定年以外のデータを保持
  const existingData = [];
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    data.forEach(row => {
      if (row[0] instanceof Date) {
        if (row[0].getFullYear() !== year) {
          existingData.push(row);
        }
      }
    });
  }

  // Google Calendarから祝日を取得
  const holidays = fetchHolidays_(year);
  const newData = holidays.map(h => [h.date, h.name]);

  // 全データを結合して日付順にソート
  const allData = existingData.concat(newData);
  allData.sort((a, b) => a[0] - b[0]);

  // シートをクリアして書き込み
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }
  if (allData.length > 0) {
    sheet.getRange(2, 1, allData.length, 2).setValues(allData);
    sheet.getRange(2, 1, allData.length, 1).setNumberFormat('yyyy/mm/dd');
  }

  return allData.length;
}

// =====================================================================
//  祝日更新（メニューから手動実行）
// =====================================================================
function updateHolidays() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const yearRes = ui.prompt('祝日の更新', '取得する年を入力してください（例: 2026）', ui.ButtonSet.OK_CANCEL);
  if (yearRes.getSelectedButton() !== ui.Button.OK) return;
  const year = parseInt(yearRes.getResponseText());

  if (isNaN(year) || year < 2020 || year > 2100) {
    ui.alert('エラー', '有効な年を入力してください。', ui.ButtonSet.OK);
    return;
  }

  const count = updateHolidaysSheet_(ss, year);
  ui.alert('完了', year + '年の祝日を取得しました（全' + count + '件）。', ui.ButtonSet.OK);
}

// =====================================================================
//  祝日トリガー（毎年1月に自動実行）
// =====================================================================
function autoUpdateHolidays() {
  const now = new Date();
  if (now.getMonth() !== 0) return; // 1月以外はスキップ

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const year = now.getFullYear();
  updateHolidaysSheet_(ss, year);
}

function setupHolidayTrigger() {
  // 既存の同名トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoUpdateHolidays') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎月1日 3:00〜4:00 に実行（1月のみ実際に処理）
  ScriptApp.newTrigger('autoUpdateHolidays')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();
}

// =====================================================================
//  スタッフ設定の読み書き
// =====================================================================
function getStaffSetting_(ss, staffName) {
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return DEFAULT_CONTRACTED_HOURS;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (const row of data) {
    if (row[0] === staffName) return row[1] || DEFAULT_CONTRACTED_HOURS;
  }
  return DEFAULT_CONTRACTED_HOURS;
}

function setStaffSetting_(ss, staffName, contractedHours) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) sheet = createSettingsSheet_(ss);

  // 既存の行を探す
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === staffName) {
        sheet.getRange(i + 2, 2).setValue(contractedHours);
        return;
      }
    }
  }

  // 新規追加
  sheet.appendRow([staffName, contractedHours]);
}

// =====================================================================
//  スタッフ別シート作成（16日〜翌15日フォーマット）
// =====================================================================
function createStaffSheet_(ss, staffName, year, month) {
  if (ss.getSheetByName(staffName)) return;

  const sheet = ss.insertSheet(staffName);
  const log = LOG_SHEET_NAME;

  // --- 行1: タイトル ---
  sheet.getRange('A1').setValue(staffName + 'さんの月間勤怠一覧')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange('A1:I1').merge();

  // --- 行2: 年月・定時設定 ---
  sheet.getRange('B2').setValue(year).setFontWeight('bold');
  sheet.getRange('C2').setValue('年');
  sheet.getRange('D2').setValue(month).setFontWeight('bold');
  sheet.getRange('E2').setValue('月');
  sheet.getRange('H2').setValue('定時：');
  sheet.getRange('I2').setFormula(
    `=IFERROR(INDEX(FILTER('${SETTINGS_SHEET_NAME}'!B:B,'${SETTINGS_SHEET_NAME}'!A:A="${staffName}"),1),${DEFAULT_CONTRACTED_HOURS})`
  );

  // --- 行3: カラムヘッダー ---
  const headers = ['日', '曜日', '開始時間', '終了時間', '休憩時間', '定時内時間', '時間外', '深夜・休日', '備考'];
  sheet.getRange('A3:I3').setValues([headers]);
  sheet.getRange('A3:I3').setFontWeight('bold').setBackground('#e2e8f0')
    .setHorizontalAlignment('center');

  // --- 行4: 上部合計行 ---
  sheet.getRange('E4').setValue('合計').setFontWeight('bold').setHorizontalAlignment('center');
  // 合計の数式は下部合計を参照（行36）
  sheet.getRange('F4').setFormula('=F39').setNumberFormat('[h]:mm');
  // 上部合計行の注記なし（1分単位合計）
  sheet.getRange('A4:I4').setBackground('#f8f9fa');

  // --- 行5〜35: 日別データ（31行分、16日〜翌15日） ---
  const hol = HOLIDAYS_SHEET_NAME;
  for (let i = 0; i < 31; i++) {
    const row = i + 5;
    const dayOffset = i; // 0=16日目, 1=17日目, ...

    // A列: 日付
    // 前月16日からの日付を計算: DATE(year, month-1, 16+dayOffset)
    // 月の最終日（翌月15日）を超えたら空白
    sheet.getRange(row, 1).setFormula(
      `=LET(startDate,DATE(B2,D2-1,16),d,startDate+${dayOffset},endDate,DATE(B2,D2,${CUTOFF_DAY}),IF(d<=endDate,d,""))`
    );
    // 日付表示: 月が変わったら "m/d"、同じ月なら "d" のみ
    // → カスタム表示は数式で制御（表示用にTEXTで整形）
    sheet.getRange(row, 1).setNumberFormat('m/d');

    // B列: 曜日
    sheet.getRange(row, 2).setFormula(
      `=IF(A${row}="","",CHOOSE(WEEKDAY(A${row}),"日","月","火","水","木","金","土"))`
    );

    // C列: 開始時間（打刻ログから取得）
    sheet.getRange(row, 3).setFormula(
      `=IFERROR(INDEX(FILTER('${log}'!$A:$A,'${log}'!$D:$D=A${row},'${log}'!$B:$B="${staffName}",'${log}'!$C:$C="入室"),1),"")`
    );
    sheet.getRange(row, 3).setNumberFormat('H:mm');

    // D列: 終了時間（打刻ログから取得）
    sheet.getRange(row, 4).setFormula(
      `=IFERROR(INDEX(FILTER('${log}'!$A:$A,'${log}'!$D:$D=A${row},'${log}'!$B:$B="${staffName}",'${log}'!$C:$C="退室"),1),"")`
    );
    sheet.getRange(row, 4).setNumberFormat('H:mm');

    // E列: 休憩時間（13時以降の退勤なら1時間、13時前なら0）
    sheet.getRange(row, 5).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),IF(HOUR(D${row})>=13,TIME(${BREAK_HOURS},0,0),""),"")`
    );
    sheet.getRange(row, 5).setNumberFormat('H:mm');

    // 休日判定: 土日 or 祝日
    const isHoliday = `OR(WEEKDAY(A${row})=1,WEEKDAY(A${row})=7,COUNTIF('${hol}'!$A:$A,A${row})>0)`;

    // F列: 定時内時間 = MIN(実働時間, 定時)（休日は0 → 全時間を深夜残業へ）
    sheet.getRange(row, 6).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),IF(${isHoliday},"",MIN(D${row}-C${row}-E${row}, $I$2/24)),"")`
    );
    sheet.getRange(row, 6).setNumberFormat('H:mm');

    // G列: 残業時間
    // 7.5h定時: 定時超〜8h（法定内残業）、8h定時: 8h超で22時前まで（通常残業）
    // 休日は0
    sheet.getRange(row, 7).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),` +
      `IF(${isHoliday},"",` +
      `IF($I$2<8,` +
      `MIN(MAX(D${row}-C${row}-E${row}-$I$2/24,0),(8-$I$2)/24),` +
      `MAX(MAX(D${row}-C${row}-E${row}-$I$2/24,0)-MIN(IF(HOUR(D${row})>=22,MOD(D${row},1)-TIME(22,0,0),0),MAX(D${row}-C${row}-E${row}-$I$2/24,0)),0)` +
      `)),"")`
    );
    sheet.getRange(row, 7).setNumberFormat('[>0]H:mm;""');

    // H列: 深夜・休日
    // 7.5h定時: 8h超（法定外残業）、8h定時: 22時以降の勤務
    // 休日は全労働時間（休日残業）
    sheet.getRange(row, 8).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),` +
      `IF(${isHoliday},D${row}-C${row}-E${row},` +
      `IF($I$2<8,` +
      `MAX(D${row}-C${row}-E${row}-8/24,0),` +
      `MIN(IF(HOUR(D${row})>=22,MOD(D${row},1)-TIME(22,0,0),0),MAX(D${row}-C${row}-E${row}-$I$2/24,0))` +
      `)),"")`
    );
    sheet.getRange(row, 8).setNumberFormat('[>0]H:mm;""');

    // I列: 備考（手動入力用）

    // J列: 祝日名（条件付き書式・備考用ヘルパー、非表示列）
    sheet.getRange(row, 10).setFormula(
      `=IFERROR(INDEX('${HOLIDAYS_SHEET_NAME}'!$B:$B,MATCH(A${row},'${HOLIDAYS_SHEET_NAME}'!$A:$A,0)),"")`
    );
  }

  // J列を非表示
  sheet.hideColumns(10);

  // --- 色定義（総合計と参照元を同色で対応） ---
  const colF = '#dce6f1'; // 定時内 = 薄い青
  const colG = '#e2efda'; // 残業 = 薄い緑
  const colH = '#fce4d6'; // 深夜残業計 = 薄いオレンジ

  // --- 行36: 下部合計行 ---
  sheet.getRange('D36').setValue('合計').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E36').setFormula('=SUMPRODUCT((E5:E35<>"")*E5:E35)');
  sheet.getRange('E36').setNumberFormat('[h]:mm').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('F36').setFormula('=SUMPRODUCT((F5:F35<>"")*F5:F35)');
  sheet.getRange('F36').setNumberFormat('[h]:mm').setFontWeight('bold').setBackground(colF);
  sheet.getRange('G36').setFormula('=SUMPRODUCT((G5:G35<>"")*G5:G35)');
  sheet.getRange('G36').setNumberFormat('[h]:mm').setFontWeight('bold').setBackground(colG);
  sheet.getRange('H36').setFormula('=SUMPRODUCT((H5:H35<>"")*H5:H35)');
  sheet.getRange('H36').setNumberFormat('[h]:mm').setFontWeight('bold').setBackground(colH);
  sheet.getRange('A36:D36').setBackground('#f8f9fa');
  sheet.getRange('I36').setBackground('#f8f9fa');
  sheet.getRange('I36').setValue('※合計は1分単位').setFontSize(8).setFontColor('#888888').setFontWeight('normal');

  // --- 行37: 深夜残業の内訳（深夜分） ---
  sheet.getRange('I37').setValue('（深夜分）').setFontSize(8).setHorizontalAlignment('left').setFontColor('#555555');
  sheet.getRange('H37').setFormula('=H36-H38');
  sheet.getRange('H37').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('A37:I37').setBackground('#f8f9fa');

  // --- 行38: 深夜残業の内訳（祝祭日分） ---
  sheet.getRange('I38').setValue('（祝祭日分）').setFontSize(8).setHorizontalAlignment('left').setFontColor('#555555');
  // 土日 or 祝日の行のH列を合計
  sheet.getRange('H38').setFormula(
    '=SUMPRODUCT(((WEEKDAY(A5:A35)=1)+(WEEKDAY(A5:A35)=7)+(J5:J35<>"")>0)*(H5:H35<>"")*H5:H35)'
  );
  sheet.getRange('H38').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('A38:I38').setBackground('#f8f9fa');

  // --- 行39: 総合計（F+G+H の合算） ---
  sheet.getRange('D39').setValue('総合計').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E39').setFormula('=E36');
  sheet.getRange('E39').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('F39').setFormula('=F36+G36+H36');
  sheet.getRange('F39').setNumberFormat('[h]:mm').setFontWeight('bold');
  // 総合計セルに参照元と同じ色を混ぜた背景（3色の参照がわかるように）
  sheet.getRange('F36').setBackground(colF);
  sheet.getRange('G36').setBackground(colG);
  sheet.getRange('H36').setBackground(colH);
  sheet.getRange('F39').setBackground('#d9d2e9'); // 薄い紫（F+G+H合算を示す）
  sheet.getRange('A39:D39').setBackground('#f8f9fa');
  sheet.getRange('G39:I39').setBackground('#f8f9fa');
  // 注記
  sheet.getRange('F40').setValue('※休憩時間は含まない。').setFontSize(8).setFontColor('#888888');

  // --- 行42: 給与計算用（10進法変換） ---
  sheet.getRange('E42').setValue('給与計算用').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(9);
  sheet.getRange('F42').setFormula('=TRUNC(F36*24,2)');
  sheet.getRange('F42').setNumberFormat('0.00').setFontWeight('bold').setBackground(colF);
  sheet.getRange('G42').setFormula('=TRUNC(G36*24,2)');
  sheet.getRange('G42').setNumberFormat('0.00').setFontWeight('bold').setBackground(colG);
  sheet.getRange('H42').setFormula('=TRUNC(H36*24,2)');
  sheet.getRange('H42').setNumberFormat('0.00').setFontWeight('bold').setBackground(colH);
  sheet.getRange('A42:E42').setBackground('#f8f9fa');
  sheet.getRange('I42').setBackground('#f8f9fa');

  // --- 行41: ↓矢印 / 7.5h定時の場合は注記 ---
  sheet.getRange('F41').setValue('↓').setHorizontalAlignment('center').setFontColor('#888888');
  const hours = getStaffSetting_(ss, staffName);
  if (hours < 8) {
    sheet.getRange('G41').setValue('（法定内）').setHorizontalAlignment('center').setFontColor('#000000').setFontSize(9);
    sheet.getRange('H41').setValue('（時間外）').setHorizontalAlignment('center').setFontColor('#000000').setFontSize(9);
  } else {
    sheet.getRange('G41').setValue('↓').setHorizontalAlignment('center').setFontColor('#888888');
    sheet.getRange('H41').setValue('↓').setHorizontalAlignment('center').setFontColor('#888888');
  }

  // --- 行43: 給与計算用 深夜内訳（深夜分） ---
  sheet.getRange('I43').setValue('（深夜分）').setFontSize(8).setHorizontalAlignment('left').setFontColor('#555555');
  sheet.getRange('H43').setFormula('=TRUNC(H37*24,2)');
  sheet.getRange('H43').setNumberFormat('0.00').setFontWeight('bold');
  sheet.getRange('A43:I43').setBackground('#f8f9fa');

  // --- 行44: 給与計算用 深夜内訳（祝祭日分） ---
  sheet.getRange('I44').setValue('（祝祭日分）').setFontSize(8).setHorizontalAlignment('left').setFontColor('#555555');
  sheet.getRange('H44').setFormula('=TRUNC(H38*24,2)');
  sheet.getRange('H44').setNumberFormat('0.00').setFontWeight('bold');
  sheet.getRange('A44:I44').setBackground('#f8f9fa');

  // --- 行45: 計算式の説明 ---
  sheet.getRange('E45').setValue('給与計算には１０進法へ変換する。計算式: 時間＋（分÷60）  小数点第3位切り捨て').setFontSize(8).setFontColor('#888888');
  sheet.getRange('E45:I45').merge();

  // --- 行47-48: 出勤日数・有給日数・欠勤日数・特別休暇（横並び） ---
  sheet.getRange('C47').setValue('出勤日数').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('D47').setValue('有給日数').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E47').setValue('欠勤日数').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('F47').setValue('特別休暇').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  // 行48: 入力セル（出勤日数は自動計算、他は手入力）
  sheet.getRange('C48').setFormula('=COUNTIFS(C5:C35,">0",D5:D35,">0")');
  sheet.getRange('C48').setHorizontalAlignment('center').setFontWeight('bold');
  // D48: 有給日数（打刻ログから自動集計）
  sheet.getRange('D48').setFormula(
    `=COUNTIFS('${LOG_SHEET_NAME}'!$B:$B,"${staffName}",'${LOG_SHEET_NAME}'!$C:$C,"有給",'${LOG_SHEET_NAME}'!$D:$D,">="&DATE(B2,D2-1,16),'${LOG_SHEET_NAME}'!$D:$D,"<="&DATE(B2,D2,${CUTOFF_DAY}))`
  );
  sheet.getRange('D48').setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange('E48:F48').setHorizontalAlignment('center').setFontWeight('bold');
  // 行48の高さを2倍に（手書き用）
  sheet.setRowHeight(48, 42);

  // --- 列幅 ---
  [50, 40, 80, 80, 80, 80, 80, 90, 160].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // --- 罫線 ---
  applyBorders_(sheet);

  // --- 条件付き書式: 土日・祝日の色分け ---
  const dataRange = sheet.getRange('A5:I35');

  const sundayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A5<>"",WEEKDAY($A5)=1)')
    .setFontColor('#dc2626')
    .setBackground('#ffe0e0')
    .setRanges([dataRange])
    .build();

  const saturdayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A5<>"",WEEKDAY($A5)=7)')
    .setFontColor('#2563eb')
    .setBackground('#e0f0ff')
    .setRanges([dataRange])
    .build();

  const holidayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A5<>"",$J5<>"")')
    .setFontColor('#dc2626')
    .setBackground('#ffe0e0')
    .setRanges([dataRange])
    .build();

  sheet.setConditionalFormatRules([sundayRule, saturdayRule, holidayRule]);

  // ヘッダー固定
  sheet.setFrozenRows(4);

  // 備考欄に祝日名を記入
  fillHolidayNames_(sheet);

  return sheet;
}

// =====================================================================
//  備考欄に祝日名を自動記入（空セルのみ）
// =====================================================================
function fillHolidayNames_(sheet) {
  SpreadsheetApp.flush(); // J列の数式を確定
  const jValues = sheet.getRange('J5:J35').getValues();  // 祝日名
  const iValues = sheet.getRange('I5:I35').getValues();  // 現在の備考

  let updated = false;
  for (let i = 0; i < 31; i++) {
    const holidayName = jValues[i][0];
    const currentNote = iValues[i][0];
    if (holidayName && (!currentNote || currentNote === '')) {
      iValues[i][0] = holidayName;
      updated = true;
    }
  }
  if (updated) {
    sheet.getRange('I5:I35').setValues(iValues);
  }
}

// =====================================================================
//  罫線の適用
// =====================================================================
function applyBorders_(sheet) {
  const border = SpreadsheetApp.BorderStyle.SOLID;
  const thinColor = '#999999';
  const headerColor = '#333333';

  // ヘッダー行（行3）: 下線を太め
  sheet.getRange('A3:I3').setBorder(true, true, true, true, true, null, headerColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 上部合計行（行4）: 全罫線
  sheet.getRange('A4:I4').setBorder(true, true, true, true, true, null, thinColor, border);

  // データ行（5〜35）: 全セルに細罫線
  sheet.getRange('A5:I35').setBorder(true, true, true, true, true, true, thinColor, border);

  // 下部合計行（行36）: 上線を太め + 全罫線
  sheet.getRange('A36:I36').setBorder(true, true, true, true, true, null, headerColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // G36:H36の下罫線を太線に
  sheet.getRange('G36:H36').setBorder(null, null, true, null, null, null, headerColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 深夜内訳行（行37-38）: H列に細罫線
  sheet.getRange('H37:H38').setBorder(true, true, true, true, true, null, thinColor, border);

  // 総合計行（行39）: 罫線で囲む
  sheet.getRange('D39:F39').setBorder(true, true, true, true, true, null, headerColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 外枠全体（A3:I39）を太めの枠で囲む
  sheet.getRange('A3:I39').setBorder(true, true, true, true, null, null, headerColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 給与計算用（行42）: 太線で囲む
  sheet.getRange('E42:H42').setBorder(true, true, true, true, true, null, headerColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // G42:H42の下罫線を太線に
  sheet.getRange('G42:H42').setBorder(null, null, true, null, null, null, headerColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // 深夜内訳（行43-44）: H列に細罫線
  sheet.getRange('H43:H44').setBorder(true, true, true, true, true, null, thinColor, border);

  // 出勤日数等の項目・入力欄（C47:F48）を罫線で囲む
  sheet.getRange('C47:F48').setBorder(true, true, true, true, true, true, thinColor, border);
}

// =====================================================================
//  既存シートへ罫線を一括適用（メニューから実行）
// =====================================================================
function applyBordersToAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '罫線の適用',
    '全スタッフシートに罫線を適用します。\n続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  const sheets = ss.getSheets();
  let count = 0;
  sheets.forEach(s => {
    if (SYSTEM_SHEET_NAMES.includes(s.getName())) return;
    applyBorders_(s);
    count++;
  });

  ui.alert('完了', count + '名分のシートに罫線を適用しました。', ui.ButtonSet.OK);
}

// =====================================================================
//  既存シートの数式を一括更新（メニューから実行）
// =====================================================================
function updateAllFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '数式の一括更新',
    '全スタッフシートのE〜H列（休憩・定時内・残業・深夜残業）の数式を最新に更新します。\n' +
    '勤務日数・有給日数の欄も追加されます。\n続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  const sheets = ss.getSheets();
  let count = 0;

  sheets.forEach(s => {
    const name = s.getName();
    if (SYSTEM_SHEET_NAMES.includes(name)) return;
    if (s.isSheetHidden()) return;

    updateSheetFormulas_(s, name);
    count++;
  });

  ui.alert('完了', count + '名分のシートの数式を更新しました。', ui.ButtonSet.OK);
}

// =====================================================================
//  個別スタッフの数式を更新（メニューから実行）
// =====================================================================
function updateSingleStaffFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const staffSheets = ss.getSheets()
    .filter(s => !SYSTEM_SHEET_NAMES.includes(s.getName()) && !s.isSheetHidden())
    .map(s => s.getName());

  if (staffSheets.length === 0) {
    ui.alert('エラー', 'スタッフシートが見つかりません。', ui.ButtonSet.OK);
    return;
  }

  const res = ui.prompt(
    '個別スタッフの数式更新',
    'スタッフ名を入力してください:\n\n' + staffSheets.join('、'),
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const name = res.getResponseText().trim();
  const sheet = ss.getSheetByName(name);
  if (!sheet || SYSTEM_SHEET_NAMES.includes(name)) {
    ui.alert('エラー', '「' + name + '」というスタッフシートが見つかりません。', ui.ButtonSet.OK);
    return;
  }

  updateSheetFormulas_(sheet, name);
  ui.alert('完了', name + 'さんのシートの数式を更新しました。', ui.ButtonSet.OK);
}

function updateSheetFormulas_(sheet, staffName) {
  const log = LOG_SHEET_NAME;
  const hol = HOLIDAYS_SHEET_NAME;

  // --- 行2: レイアウト移行（旧: C2=年：,D2=年,E2=月：,F2=月 → 新: B2=年,C2=年,D2=月,E2=月） ---
  const oldD2 = sheet.getRange('D2').getValue();
  const oldF2 = sheet.getRange('F2').getValue();
  const oldC2 = sheet.getRange('C2').getValue();
  if (String(oldC2).indexOf('：') !== -1 && oldF2) {
    // 旧レイアウト検出: C2が「年：」（コロン付き）でF2に月の値がある
    sheet.getRange('B2').setValue(oldD2).setFontWeight('bold');
    sheet.getRange('C2').setValue('年');
    sheet.getRange('D2').setValue(oldF2).setFontWeight('bold');
    sheet.getRange('E2').setValue('月');
    sheet.getRange('F2').clearContent();
  }

  // A列・B列の数式を更新（B2/D2参照に統一）
  for (let i = 0; i < 31; i++) {
    const row = i + 5;
    const dayOffset = i;
    sheet.getRange(row, 1).setFormula(
      `=LET(startDate,DATE(B2,D2-1,16),d,startDate+${dayOffset},endDate,DATE(B2,D2,${CUTOFF_DAY}),IF(d<=endDate,d,""))`
    );
    sheet.getRange(row, 2).setFormula(
      `=IF(A${row}="","",CHOOSE(WEEKDAY(A${row}),"日","月","火","水","木","金","土"))`
    );
  }

  // C列・D列: 開始/終了時間のFILTER数式を復元（手入力で上書きされた場合の修復）
  for (let i = 0; i < 31; i++) {
    const row = i + 5;
    sheet.getRange(row, 3).setFormula(
      `=IFERROR(INDEX(FILTER('${log}'!$A:$A,'${log}'!$D:$D=A${row},'${log}'!$B:$B="${staffName}",'${log}'!$C:$C="入室"),1),"")`
    );
    sheet.getRange(row, 4).setFormula(
      `=IFERROR(INDEX(FILTER('${log}'!$A:$A,'${log}'!$D:$D=A${row},'${log}'!$B:$B="${staffName}",'${log}'!$C:$C="退室"),1),"")`
    );
  }
  sheet.getRange(5, 3, 31, 1).setNumberFormat('H:mm');
  sheet.getRange(5, 4, 31, 1).setNumberFormat('H:mm');

  // E〜H列の数式を配列で一括構築（行5〜35）
  const formulas = [];
  for (let i = 0; i < 31; i++) {
    const row = i + 5;
    // 休日判定: 土日 or 祝日
    const isHoliday = `OR(WEEKDAY(A${row})=1,WEEKDAY(A${row})=7,COUNTIF('${hol}'!$A:$A,A${row})>0)`;
    formulas.push([
      // E列: 休憩（13時前退勤なら0）
      `=IF(AND(C${row}<>"",D${row}<>""),IF(HOUR(D${row})>=13,TIME(${BREAK_HOURS},0,0),""),"")`,
      // F列: 定時内（休日は空）
      `=IF(AND(C${row}<>"",D${row}<>""),IF(${isHoliday},"",MIN(D${row}-C${row}-E${row}, $I$2/24)),"")`,
      // G列: 残業（7.5h:法定内、8h:22時前の通常残業、休日は空）
      `=IF(AND(C${row}<>"",D${row}<>""),IF(${isHoliday},"",IF($I$2<8,MIN(MAX(D${row}-C${row}-E${row}-$I$2/24,0),(8-$I$2)/24),MAX(MAX(D${row}-C${row}-E${row}-$I$2/24,0)-MIN(IF(HOUR(D${row})>=22,MOD(D${row},1)-TIME(22,0,0),0),MAX(D${row}-C${row}-E${row}-$I$2/24,0)),0))),"")`,
      // H列: 深夜残業（7.5h:8h超法定外、8h:22時以降、休日は全労働時間）
      `=IF(AND(C${row}<>"",D${row}<>""),IF(${isHoliday},D${row}-C${row}-E${row},IF($I$2<8,MAX(D${row}-C${row}-E${row}-8/24,0),MIN(IF(HOUR(D${row})>=22,MOD(D${row},1)-TIME(22,0,0),0),MAX(D${row}-C${row}-E${row}-$I$2/24,0)))),"")`,
    ]);
  }

  // E5:H35に一括書き込み
  sheet.getRange(5, 5, 31, 4).setFormulas(formulas);

  // J列: 祝日名（条件付き書式・備考用ヘルパー）
  const jFormulas = [];
  for (let i = 0; i < 31; i++) {
    const row = i + 5;
    jFormulas.push([`=IFERROR(INDEX('${hol}'!$B:$B,MATCH(A${row},'${hol}'!$A:$A,0)),"")`]);
  }
  sheet.getRange(5, 10, 31, 1).setFormulas(jFormulas);
  sheet.hideColumns(10);

  // フォーマット一括設定
  sheet.getRange(5, 5, 31, 1).setNumberFormat('H:mm');
  sheet.getRange(5, 6, 31, 1).setNumberFormat('H:mm');
  sheet.getRange(5, 7, 31, 1).setNumberFormat('[>0]H:mm;""');
  sheet.getRange(5, 8, 31, 1).setNumberFormat('[>0]H:mm;""');

  // カラムヘッダーを更新
  sheet.getRange('G3').setValue('時間外');
  sheet.getRange('H3').setValue('深夜・休日');

  // 旧レイアウト（行36〜48）をクリア（行36の数式は残す、書式のみリセット）
  sheet.getRange('A37:I48').clearContent().clearFormat();
  try { sheet.getRange('E45:I45').breakApart(); } catch(e) {}
  try { sheet.getRange('E41:I41').breakApart(); } catch(e) {}

  // --- 色定義 ---
  const colF = '#dce6f1';
  const colG = '#e2efda';
  const colH = '#fce4d6';

  // 行36: 合計行
  sheet.getRange('D36').setValue('合計').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E36').setFormula('=SUMPRODUCT((E5:E35<>"")*E5:E35)');
  sheet.getRange('E36').setNumberFormat('[h]:mm').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('F36').setBackground(colF);
  sheet.getRange('G36').setBackground(colG);
  sheet.getRange('H36').setBackground(colH);
  sheet.getRange('A36:D36').setBackground('#f8f9fa');
  sheet.getRange('I36').setBackground('#f8f9fa');
  sheet.getRange('I36').setValue('※合計は1分単位').setFontSize(8).setFontColor('#888888').setFontWeight('normal');

  // 行37: 深夜内訳（深夜分）
  sheet.getRange('I37').setValue('（深夜分）').setFontSize(8).setHorizontalAlignment('left').setFontColor('#555555');
  sheet.getRange('H37').setFormula('=H36-H38');
  sheet.getRange('H37').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('A37:I37').setBackground('#f8f9fa');

  // 行38: 深夜内訳（祝祭日分）
  sheet.getRange('I38').setValue('（祝祭日分）').setFontSize(8).setHorizontalAlignment('left').setFontColor('#555555');
  sheet.getRange('H38').setFormula(
    '=SUMPRODUCT(((WEEKDAY(A5:A35)=1)+(WEEKDAY(A5:A35)=7)+(J5:J35<>"")>0)*(H5:H35<>"")*H5:H35)'
  );
  sheet.getRange('H38').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('A38:I38').setBackground('#f8f9fa');

  // 行39: 総合計
  sheet.getRange('D39').setValue('総合計').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E39').setFormula('=E36');
  sheet.getRange('E39').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('F39').setFormula('=F36+G36+H36');
  sheet.getRange('F39').setNumberFormat('[h]:mm').setFontWeight('bold').setBackground('#d9d2e9');
  sheet.getRange('A39:D39').setBackground('#f8f9fa');
  sheet.getRange('G39:I39').setBackground('#f8f9fa');
  // 注記
  sheet.getRange('F40').setValue('※休憩時間は含まない。').setFontSize(8).setFontColor('#888888');

  // 行41: ↓矢印 / 7.5h定時の場合は注記
  sheet.getRange('F41').setValue('↓').setHorizontalAlignment('center').setFontColor('#888888');
  const ss = sheet.getParent();
  const hours = getStaffSetting_(ss, staffName);
  if (hours < 8) {
    sheet.getRange('G41').setValue('（法定内）').setHorizontalAlignment('center').setFontColor('#000000').setFontSize(9);
    sheet.getRange('H41').setValue('（時間外）').setHorizontalAlignment('center').setFontColor('#000000').setFontSize(9);
  } else {
    sheet.getRange('G41').setValue('↓').setHorizontalAlignment('center').setFontColor('#888888');
    sheet.getRange('H41').setValue('↓').setHorizontalAlignment('center').setFontColor('#888888');
  }

  // 行42: 給与計算用（10進法）
  sheet.getRange('E42').setValue('給与計算用').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(9);
  sheet.getRange('F42').setFormula('=TRUNC(F36*24,2)');
  sheet.getRange('F42').setNumberFormat('0.00').setFontWeight('bold').setBackground(colF);
  sheet.getRange('G42').setFormula('=TRUNC(G36*24,2)');
  sheet.getRange('G42').setNumberFormat('0.00').setFontWeight('bold').setBackground(colG);
  sheet.getRange('H42').setFormula('=TRUNC(H36*24,2)');
  sheet.getRange('H42').setNumberFormat('0.00').setFontWeight('bold').setBackground(colH);
  sheet.getRange('A42:E42').setBackground('#f8f9fa');
  sheet.getRange('I42').setBackground('#f8f9fa');

  // 行43: 給与計算用 深夜内訳（深夜分）
  sheet.getRange('I43').setValue('（深夜分）').setFontSize(8).setHorizontalAlignment('left').setFontColor('#555555');
  sheet.getRange('H43').setFormula('=TRUNC(H37*24,2)');
  sheet.getRange('H43').setNumberFormat('0.00').setFontWeight('bold');
  sheet.getRange('A43:I43').setBackground('#f8f9fa');

  // 行44: 給与計算用 深夜内訳（祝祭日分）
  sheet.getRange('I44').setValue('（祝祭日分）').setFontSize(8).setHorizontalAlignment('left').setFontColor('#555555');
  sheet.getRange('H44').setFormula('=TRUNC(H38*24,2)');
  sheet.getRange('H44').setNumberFormat('0.00').setFontWeight('bold');
  sheet.getRange('A44:I44').setBackground('#f8f9fa');

  // 行45: 計算式の説明
  sheet.getRange('E45').setValue('給与計算には１０進法へ変換する。計算式: 時間＋（分÷60）  小数点第3位切り捨て').setFontSize(8).setFontColor('#888888');
  sheet.getRange('E45:I45').merge();

  // 行47-48: 出勤日数・有給日数・欠勤日数・特別休暇（横並び）
  sheet.getRange('C47').setValue('出勤日数').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('D47').setValue('有給日数').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E47').setValue('欠勤日数').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('F47').setValue('特別休暇').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  // 行48: 入力セル（出勤日数は自動計算、他は手入力）
  sheet.getRange('C48').setFormula('=COUNTIFS(C5:C35,">0",D5:D35,">0")');
  sheet.getRange('C48').setHorizontalAlignment('center').setFontWeight('bold');
  // D48: 有給日数（打刻ログから自動集計）
  sheet.getRange('D48').setFormula(
    `=COUNTIFS('${LOG_SHEET_NAME}'!$B:$B,"${staffName}",'${LOG_SHEET_NAME}'!$C:$C,"有給",'${LOG_SHEET_NAME}'!$D:$D,">="&DATE(B2,D2-1,16),'${LOG_SHEET_NAME}'!$D:$D,"<="&DATE(B2,D2,${CUTOFF_DAY}))`
  );
  sheet.getRange('D48').setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange('E48:F48').setHorizontalAlignment('center').setFontWeight('bold');
  // 行48の高さを2倍に（手書き用）
  sheet.setRowHeight(48, 42);

  // F列幅をC列と統一（80px）
  sheet.setColumnWidth(6, 80);

  // 罫線も適用
  applyBorders_(sheet);

  // 備考欄に祝日名を記入（空セルのみ）
  fillHolidayNames_(sheet);

  // 条件付き書式を更新（土日 + 祝日）
  const dataRange = sheet.getRange('A5:I35');

  const sundayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A5<>"",WEEKDAY($A5)=1)')
    .setFontColor('#dc2626')
    .setBackground('#ffe0e0')
    .setRanges([dataRange])
    .build();

  const saturdayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A5<>"",WEEKDAY($A5)=7)')
    .setFontColor('#2563eb')
    .setBackground('#e0f0ff')
    .setRanges([dataRange])
    .build();

  const holidayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A5<>"",$J5<>"")')
    .setFontColor('#dc2626')
    .setBackground('#ffe0e0')
    .setRanges([dataRange])
    .build();

  sheet.setConditionalFormatRules([sundayRule, saturdayRule, holidayRule]);
}

// =====================================================================
//  打刻処理（HTML画面から呼ばれる）
// =====================================================================
function recordClock(staffName, type) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) throw new Error('打刻ログシートが見つかりません。初期設定を実行してください。');

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // 重複チェック: 同じ日・同じ人・同じ種別があればエラー
  if (logSheet.getLastRow() > 1) {
    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();
    for (const row of data) {
      if (!row[0]) continue;
      const rowDate = new Date(row[3]);
      if (rowDate.getTime() === today.getTime() && row[1] === staffName && row[2] === type) {
        throw new Error(staffName + 'さんは本日すでに' + type + '打刻済みです');
      }
    }
  }

  // 打刻ログに追記
  logSheet.appendRow([now, staffName, type, today]);

  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow, 1).setNumberFormat('yyyy/MM/dd HH:mm');
  logSheet.getRange(lastRow, 4).setNumberFormat('yyyy/MM/dd');

  return {
    success: true,
    message: staffName + 'さんの' + type + 'を記録しました',
    time: Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm')
  };
}

// =====================================================================
//  有給申請（HTML画面から呼ばれる）
// =====================================================================
function recordPaidLeave(staffName, targetDateStr, memo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) throw new Error('打刻ログシートが見つかりません。初期設定を実行してください。');

  // 日付文字列をパース（"2026-04-16" 形式）
  const parts = targetDateStr.split('-');
  const targetDate = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  targetDate.setHours(0, 0, 0, 0);

  // 重複チェック: 同じ日・同じ人の有給がすでにあればエラー
  if (logSheet.getLastRow() > 1) {
    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();
    for (const row of data) {
      if (!row[0]) continue;
      const rowDate = new Date(row[3]);
      rowDate.setHours(0, 0, 0, 0);
      if (rowDate.getTime() === targetDate.getTime() && row[1] === staffName && row[2] === '有給') {
        const dateLabel = (targetDate.getMonth() + 1) + '/' + targetDate.getDate();
        throw new Error(staffName + 'さんの' + dateLabel + 'の有給はすでに申請済みです');
      }
    }
  }

  // 打刻ログに追記（A列=申請日時, B列=氏名, C列="有給", D列=対象日付）
  const now = new Date();
  logSheet.appendRow([now, staffName, '有給', targetDate]);

  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow, 1).setNumberFormat('yyyy/MM/dd HH:mm');
  logSheet.getRange(lastRow, 4).setNumberFormat('yyyy/MM/dd');

  // 対象日の備考列に有給を記入
  const staffSheet = ss.getSheetByName(staffName);
  if (staffSheet) {
    for (let i = 0; i < 31; i++) {
      const cellDate = staffSheet.getRange(i + 5, 1).getValue();
      if (cellDate instanceof Date) {
        const d = new Date(cellDate);
        d.setHours(0, 0, 0, 0);
        if (d.getTime() === targetDate.getTime()) {
          const noteText = (memo && memo.trim()) ? '有給（' + memo.trim() + '）' : '有給';
          staffSheet.getRange(i + 5, 9).setValue(noteText);
          break;
        }
      }
    }
  }

  const dateLabel = (targetDate.getMonth() + 1) + '/' + targetDate.getDate();
  return {
    success: true,
    message: staffName + 'さんの' + dateLabel + 'の有給を申請しました'
  };
}

// =====================================================================
//  スタッフ一覧＋本日のステータス取得（HTML画面から呼ばれる）
// =====================================================================
function getStaffWithStatus() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let todayLogs = [];
  if (logSheet && logSheet.getLastRow() > 1) {
    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();
    todayLogs = data.filter(row => {
      if (!row[3]) return false;
      const d = new Date(row[3]);
      return d.getTime() === today.getTime();
    });
  }

  const sheets = ss.getSheets();
  const staffNames = sheets
    .filter(s => !s.isSheetHidden())
    .map(s => s.getName())
    .filter(n => !SYSTEM_SHEET_NAMES.includes(n));

  // スタッフ設定シートの表示順でソート
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const sortMap = {};
  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    const cols = settingsSheet.getLastColumn();
    const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, Math.max(cols, 3)).getValues();
    data.forEach(row => {
      sortMap[row[0]] = row[2] !== '' && row[2] !== undefined ? row[2] : 9999;
    });
  }
  staffNames.sort((a, b) => ((sortMap[a] ?? 9999) - (sortMap[b] ?? 9999)));

  return staffNames.map(name => {
    let status = 'none';
    let clockIn = null;
    let clockOut = null;

    for (const row of todayLogs) {
      if (row[1] !== name) continue;
      if (row[2] === '入室') {
        clockIn = Utilities.formatDate(new Date(row[0]), 'Asia/Tokyo', 'HH:mm');
        status = 'working';
      }
      if (row[2] === '退室') {
        clockOut = Utilities.formatDate(new Date(row[0]), 'Asia/Tokyo', 'HH:mm');
        status = 'done';
      }
    }

    return { name, status, clockIn, clockOut };
  });
}

// =====================================================================
//  月を移行
// =====================================================================
function promptNewMonth() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const res = ui.prompt('月を移行', '年/月を入力してください（例: 2026/5）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const parts = res.getResponseText().split('/');
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);

  if (parts.length !== 2 || isNaN(year) || isNaN(month) || month < 1 || month > 12 || year < 2020) {
    ui.alert('エラー', '「年/月」の形式で入力してください。\n例: 2026/5', ui.ButtonSet.OK);
    return;
  }

  const confirm = ui.alert(
    '確認',
    '各スタッフのシートを ' + year + '年' + month + '月に切り替えます。\n' +
    '（勤務期間: ' + (month === 1 ? 12 : month - 1) + '/16〜' + month + '/15）\n\n' +
    '備考欄は自動バックアップ後にクリアされます。\n\n' +
    '※ 打刻ログの生データは保持されます。\n' +
    '※ 備考は「備考ログ」シートに保存され、後から復元できます。\n\n' +
    '続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  let notesLog = ss.getSheetByName(NOTES_LOG_SHEET_NAME);
  if (!notesLog) notesLog = createNotesLogSheet_(ss);

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (SYSTEM_SHEET_NAMES.includes(name)) return;

    const curYear = sheet.getRange('B2').getValue();
    const curMonth = sheet.getRange('D2').getValue();

    // 備考欄（I列）をバックアップ
    const notes = sheet.getRange('I5:I35').getValues();
    const rows = [];
    notes.forEach((cell, i) => {
      if (cell[0] !== '' && cell[0] !== null) {
        rows.push([curYear, curMonth, name, i + 1, cell[0]]);
      }
    });
    if (rows.length > 0) {
      if (notesLog.getLastRow() > 1) {
        const existing = notesLog.getRange(2, 1, notesLog.getLastRow() - 1, 5).getValues();
        for (let r = existing.length - 1; r >= 0; r--) {
          if (existing[r][0] == curYear && existing[r][1] == curMonth && existing[r][2] === name) {
            notesLog.deleteRow(r + 2);
          }
        }
      }
      notesLog.getRange(notesLog.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    }

    // 年月更新＆備考クリア
    sheet.getRange('B2').setValue(year);
    sheet.getRange('D2').setValue(month);
    sheet.getRange('I5:I35').clearContent();

    // 備考欄に祝日名を記入
    fillHolidayNames_(sheet);
  });

  ui.alert('完了', '全スタッフのシートを ' + year + '年' + month + '月に更新しました。\n' +
    '（勤務期間: ' + (month === 1 ? 12 : month - 1) + '/16〜' + month + '/15）', ui.ButtonSet.OK);
}

// =====================================================================
//  備考を復元
// =====================================================================
function restoreNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const notesLog = ss.getSheetByName(NOTES_LOG_SHEET_NAME);
  if (!notesLog || notesLog.getLastRow() <= 1) {
    ui.alert('備考ログ', '復元できる備考データがありません。', ui.ButtonSet.OK);
    return;
  }

  const staffSheets = ss.getSheets().filter(s =>
    !SYSTEM_SHEET_NAMES.includes(s.getName())
  );
  if (staffSheets.length === 0) return;

  const year = staffSheets[0].getRange('B2').getValue();
  const month = staffSheets[0].getRange('D2').getValue();

  const confirm = ui.alert(
    '備考を復元',
    year + '年' + month + '月の備考データを復元します。\n' +
    '現在の備考欄は上書きされます。続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  const logData = notesLog.getRange(2, 1, notesLog.getLastRow() - 1, 5).getValues();
  let restoredCount = 0;

  staffSheets.forEach(sheet => {
    const staffName = sheet.getName();
    const matches = logData.filter(row =>
      row[0] == year && row[1] == month && row[2] === staffName
    );
    matches.forEach(row => {
      const day = row[3]; // 日（1〜31、行番号のオフセット）
      const note = row[4];
      if (day >= 1 && day <= 31) {
        sheet.getRange(day + 4, 9).setValue(note); // I列（9列目）、行5が1番目
        restoredCount++;
      }
    });
  });

  if (restoredCount > 0) {
    ui.alert('復元完了', year + '年' + month + '月の備考を' + restoredCount + '件復元しました。', ui.ButtonSet.OK);
  } else {
    ui.alert('備考ログ', year + '年' + month + '月の備考データは見つかりませんでした。', ui.ButtonSet.OK);
  }
}

// =====================================================================
//  PDF共通ヘルパー
// =====================================================================

// フォルダ取得/作成: 勤怠管理PDF / {year} / {month(2桁)}
function getOrCreateFolder_(year, month) {
  const monthStr = String(month).padStart(2, '0');

  let rootFolder;
  const rootFolders = DriveApp.getFoldersByName(PDF_ROOT_FOLDER_NAME);
  if (rootFolders.hasNext()) {
    rootFolder = rootFolders.next();
  } else {
    rootFolder = DriveApp.createFolder(PDF_ROOT_FOLDER_NAME);
  }

  let yearFolder;
  const yearFolders = rootFolder.getFoldersByName(String(year));
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = rootFolder.createFolder(String(year));
  }

  let monthFolder;
  const monthFolders = yearFolder.getFoldersByName(monthStr);
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = yearFolder.createFolder(monthStr);
  }

  return monthFolder;
}

// 1名分のPDF生成 → ドライブに保存して結果を返す
function generateSinglePDF_(ss, sheet, year, month, folder) {
  const staffName = sheet.getName();
  const monthStr = String(month).padStart(2, '0');
  const fileName = staffName + '_' + year + '年' + monthStr + '月_tenemosタイムカード.pdf';

  // シートのB2/D2が指定年月と異なる場合、一時変更
  const origYear = sheet.getRange('B2').getValue();
  const origMonth = sheet.getRange('D2').getValue();
  const needRestore = (origYear != year || origMonth != month);

  if (needRestore) {
    sheet.getRange('B2').setValue(year);
    sheet.getRange('D2').setValue(month);
    SpreadsheetApp.flush();
  }

  // PDF出力前に罫線を確実に適用
  applyBorders_(sheet);

  // 数式の再計算を確実に完了させる（FILTER数式等の反映待ち）
  SpreadsheetApp.flush();

  // 同じスタッフ・同じ月のPDFが既にあれば削除（ファイル名変更にも対応）
  const prefix = staffName + '_' + year + '年' + monthStr + '月_';
  const allFiles = folder.getFiles();
  while (allFiles.hasNext()) {
    const f = allFiles.next();
    if (f.getName().startsWith(prefix)) {
      f.setTrashed(true);
    }
  }

  // PDF生成（A1:I36 = タイトル〜下部合計行）
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'format=pdf' +
    '&gid=' + sheet.getSheetId() +
    '&size=A4' +
    '&landscape=true' +
    '&fitw=true' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&sheetnames=false' +
    '&pagenum=UNDEFINED' +
    '&fzr=true' +
    '&range=A1:I48';

  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
  });

  const blob = response.getBlob().setName(fileName);
  const file = folder.createFile(blob);

  if (needRestore) {
    sheet.getRange('B2').setValue(origYear);
    sheet.getRange('D2').setValue(origMonth);
    SpreadsheetApp.flush();
  }

  return {
    fileId: file.getId(),
    fileName: file.getName(),
    url: file.getUrl(),
    staffName: staffName,
    year: year,
    month: month
  };
}

// =====================================================================
//  PDF出力（月次）— メニューから呼び出し
// =====================================================================
function exportPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const staffSheets = ss.getSheets().filter(s => !SYSTEM_SHEET_NAMES.includes(s.getName()));
  if (staffSheets.length === 0) {
    ui.alert('エラー', 'スタッフシートが見つかりません。先に初期設定を実行してください。', ui.ButtonSet.OK);
    return;
  }

  const year = staffSheets[0].getRange('B2').getValue();
  const month = staffSheets[0].getRange('D2').getValue();
  const staffNames = staffSheets.map(s => s.getName()).join('、');

  const confirm = ui.alert(
    'PDF出力',
    year + '年' + month + '月の勤務表をPDFで出力します。\n\n' +
    '対象スタッフ: ' + staffNames + '\n\n' +
    'Googleドライブに保存されます。続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  const folder = getOrCreateFolder_(year, month);
  const createdFiles = [];

  staffSheets.forEach(sheet => {
    const result = generateSinglePDF_(ss, sheet, year, month, folder);
    createdFiles.push(result.staffName);
  });

  ui.alert(
    'PDF出力完了',
    createdFiles.length + '名分のPDFを出力しました。\n\n' +
    '保存先: Googleドライブ\n' +
    '  ' + PDF_ROOT_FOLDER_NAME + ' / ' + year + ' / ' + String(month).padStart(2, '0') + '\n\n' +
    '出力ファイル:\n' +
    createdFiles.map(name => '  ・' + name).join('\n'),
    ui.ButtonSet.OK
  );
}

// =====================================================================
//  管理者ページ API（doPost）
// =====================================================================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);

    // APIキー認証
    const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (!apiKey || params.apiKey !== apiKey) {
      return jsonResponse_({ success: false, error: '認証エラー' });
    }

    switch (params.action) {
      case 'generatePDF':
        return jsonResponse_(apiGeneratePDF_(params));
      case 'listPDFs':
        return jsonResponse_(apiListPDFs_(params));
      case 'getPDFContent':
        return jsonResponse_(apiGetPDFContent_(params));
      case 'getStaffList':
        return jsonResponse_(apiGetStaffList_());
      case 'setStaffSetting':
        return jsonResponse_(apiSetStaffSetting_(params));
      case 'addStaff':
        return jsonResponse_(apiAddStaff_(params));
      case 'removeStaff':
        return jsonResponse_(apiRemoveStaff_(params));
      case 'updateStaffOrder':
        return jsonResponse_(apiUpdateStaffOrder_(params));
      case 'renameStaff':
        return jsonResponse_(apiRenameStaff_(params));
      case 'getClockLog':
        return jsonResponse_(apiGetClockLog_(params));
      case 'addClockEntry':
        return jsonResponse_(apiAddClockEntry_(params));
      case 'editClockEntry':
        return jsonResponse_(apiEditClockEntry_(params));
      case 'deleteClockEntry':
        return jsonResponse_(apiDeleteClockEntry_(params));
      default:
        return jsonResponse_({ success: false, error: '不明なアクション: ' + params.action });
    }
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- API: PDF生成 ---
function apiGeneratePDF_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const year = params.year;
  const month = params.month;
  const folder = getOrCreateFolder_(year, month);

  if (!params.staffName || params.staffName === 'all') {
    const staffSheets = ss.getSheets().filter(s => !SYSTEM_SHEET_NAMES.includes(s.getName()));
    const results = staffSheets.map(sheet => generateSinglePDF_(ss, sheet, year, month, folder));
    return { success: true, data: results };
  }

  const sheet = ss.getSheetByName(params.staffName);
  if (!sheet) {
    return { success: false, error: 'スタッフシートが見つかりません: ' + params.staffName };
  }
  const result = generateSinglePDF_(ss, sheet, year, month, folder);
  return { success: true, data: result };
}

// --- API: PDF一覧取得 ---
function apiListPDFs_(params) {
  const year = params.year;

  let rootFolder;
  const rootFolders = DriveApp.getFoldersByName(PDF_ROOT_FOLDER_NAME);
  if (!rootFolders.hasNext()) {
    return { success: true, data: [] };
  }
  rootFolder = rootFolders.next();

  let yearFolder;
  const yearFolders = rootFolder.getFoldersByName(String(year));
  if (!yearFolders.hasNext()) {
    return { success: true, data: [] };
  }
  yearFolder = yearFolders.next();

  const results = [];
  const filterMonth = params.month ? parseInt(params.month) : null;

  // 月指定がある場合はそのフォルダだけ、なければ全月
  const monthFoldersToScan = [];
  if (filterMonth) {
    const mStr = String(filterMonth).padStart(2, '0');
    const mfs = yearFolder.getFoldersByName(mStr);
    if (mfs.hasNext()) monthFoldersToScan.push(mfs.next());
  } else {
    const allMf = yearFolder.getFolders();
    while (allMf.hasNext()) monthFoldersToScan.push(allMf.next());
  }

  for (const mFolder of monthFoldersToScan) {
    const monthNum = parseInt(mFolder.getName());
    const files = mFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();

      if (params.staffName && params.staffName !== '' && !name.includes(params.staffName)) {
        continue;
      }

      // ファイル名からスタッフ名を抽出（例: "飯島祥子_2026年04月_勤怠表.pdf" → "飯島祥子"）
      const staffMatch = name.match(/^(.+?)_\d{4}年/);
      const staffName = staffMatch ? staffMatch[1] : '';

      results.push({
        fileId: file.getId(),
        fileName: name,
        staffName: staffName,
        year: year,
        month: monthNum,
        createdAt: file.getDateCreated().toISOString(),
        url: file.getUrl()
      });
    }
  }

  results.sort((a, b) => a.month - b.month || a.fileName.localeCompare(b.fileName));
  return { success: true, data: results };
}

// --- API: PDFコンテンツ取得（Base64） ---
function apiGetPDFContent_(params) {
  const file = DriveApp.getFileById(params.fileId);
  const blob = file.getBlob();
  return {
    success: true,
    data: {
      fileName: file.getName(),
      mimeType: blob.getContentType(),
      base64: Utilities.base64Encode(blob.getBytes())
    }
  };
}

// --- API: スタッフ一覧取得（定時設定付き・表示順ソート） ---
function apiGetStaffList_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffNames = ss.getSheets()
    .filter(s => !s.isSheetHidden())
    .map(s => s.getName())
    .filter(n => !SYSTEM_SHEET_NAMES.includes(n));

  // スタッフ設定シートから表示順を取得
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const sortMap = {};
  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    const cols = settingsSheet.getLastColumn();
    const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, Math.max(cols, 3)).getValues();
    data.forEach(row => {
      sortMap[row[0]] = { contractedHours: row[1] || DEFAULT_CONTRACTED_HOURS, sortOrder: row[2] };
    });
  }

  const staffList = staffNames.map(name => ({
    name: name,
    contractedHours: (sortMap[name] && sortMap[name].contractedHours) || DEFAULT_CONTRACTED_HOURS,
    sortOrder: (sortMap[name] && sortMap[name].sortOrder !== '' && sortMap[name].sortOrder !== undefined) ? sortMap[name].sortOrder : 9999
  }));

  staffList.sort((a, b) => a.sortOrder - b.sortOrder);

  return { success: true, data: staffList };
}

// --- API: スタッフ定時設定の更新 ---
function apiSetStaffSetting_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffName = params.staffName;
  const contractedHours = parseFloat(params.contractedHours);

  if (!staffName) {
    return { success: false, error: 'スタッフ名が指定されていません。' };
  }
  if (isNaN(contractedHours) || contractedHours < 1 || contractedHours > 24) {
    return { success: false, error: '定時は1〜24の範囲で指定してください。' };
  }

  setStaffSetting_(ss, staffName, contractedHours);
  return { success: true, data: { staffName: staffName, contractedHours: contractedHours } };
}

// --- API: スタッフ追加 ---
function apiAddStaff_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffName = (params.staffName || '').trim();
  const contractedHours = parseFloat(params.contractedHours) || DEFAULT_CONTRACTED_HOURS;

  if (!staffName) {
    return { success: false, error: 'スタッフ名を指定してください。' };
  }

  if (ss.getSheetByName(staffName)) {
    return { success: false, error: 'そのスタッフは既に存在します: ' + staffName };
  }

  // 現在の年月を既存スタッフシートから取得（なければ現在月）
  const existingSheets = ss.getSheets().filter(s => !SYSTEM_SHEET_NAMES.includes(s.getName()));
  let year, month;
  if (existingSheets.length > 0) {
    year = existingSheets[0].getRange('B2').getValue();
    month = existingSheets[0].getRange('D2').getValue();
  } else {
    const now = new Date();
    year = now.getFullYear();
    month = now.getMonth() + 1;
  }

  createStaffSheet_(ss, staffName, year, month);
  setStaffSetting_(ss, staffName, contractedHours);

  return { success: true, data: { staffName: staffName, contractedHours: contractedHours } };
}

// --- API: スタッフ削除（シートを非表示にして保持） ---
function apiRemoveStaff_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffName = (params.staffName || '').trim();

  if (!staffName) {
    return { success: false, error: 'スタッフ名を指定してください。' };
  }

  const sheet = ss.getSheetByName(staffName);
  if (!sheet) {
    return { success: false, error: 'スタッフシートが見つかりません: ' + staffName };
  }

  // シートを削除せず非表示にする（データ保全のため）
  sheet.hideSheet();

  // スタッフ設定から削除
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 1).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === staffName) {
        settingsSheet.deleteRow(i + 2);
      }
    }
  }

  return { success: true, data: { staffName: staffName, message: 'シートを非表示にしました（データは保持）' } };
}

// --- API: スタッフ表示順の更新 ---
function apiUpdateStaffOrder_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const order = params.order; // [{name: '...', sortOrder: 0}, ...]

  if (!order || !Array.isArray(order)) {
    return { success: false, error: '並び順データが不正です。' };
  }

  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) settingsSheet = createSettingsSheet_(ss);

  if (settingsSheet.getLastRow() > 1) {
    const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 1).getValues();
    for (const item of order) {
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === item.name) {
          settingsSheet.getRange(i + 2, 3).setValue(item.sortOrder);
          break;
        }
      }
    }
  }

  return { success: true };
}

// --- API: スタッフ名の変更 ---
function apiRenameStaff_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const oldName = (params.oldName || '').trim();
  const newName = (params.newName || '').trim();

  if (!oldName || !newName) {
    return { success: false, error: '変更前・変更後の氏名を指定してください。' };
  }
  if (oldName === newName) {
    return { success: false, error: '変更前と変更後の名前が同じです。' };
  }

  // 変更先の名前が既に使われていないかチェック
  if (ss.getSheetByName(newName)) {
    return { success: false, error: '「' + newName + '」という名前のシートが既に存在します。' };
  }

  // スタッフシートの存在確認
  const sheet = ss.getSheetByName(oldName);
  if (!sheet) {
    return { success: false, error: 'スタッフシートが見つかりません: ' + oldName };
  }

  // 1. シートタブ名を変更
  sheet.setName(newName);

  // 2. A1のタイトルを変更
  sheet.getRange('A1').setValue(newName + 'さんの月間勤怠一覧');

  // 3. C列・D列の数式内のスタッフ名を更新（打刻ログ参照のFILTER数式）
  for (let row = 5; row <= 35; row++) {
    const cFormula = sheet.getRange(row, 3).getFormula();
    if (cFormula) {
      sheet.getRange(row, 3).setFormula(cFormula.replace(new RegExp(escapeRegex_(oldName), 'g'), newName));
    }
    const dFormula = sheet.getRange(row, 4).getFormula();
    if (dFormula) {
      sheet.getRange(row, 4).setFormula(dFormula.replace(new RegExp(escapeRegex_(oldName), 'g'), newName));
    }
  }

  // 4. I2の定時参照数式を更新
  sheet.getRange('I2').setFormula(
    `=IFERROR(INDEX(FILTER('${SETTINGS_SHEET_NAME}'!B:B,'${SETTINGS_SHEET_NAME}'!A:A="${newName}"),1),${DEFAULT_CONTRACTED_HOURS})`
  );

  // 5. スタッフ設定シートの名前を更新
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === oldName) {
        settingsSheet.getRange(i + 2, 1).setValue(newName);
        break;
      }
    }
  }

  // 6. 打刻ログの氏名を更新
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (logSheet && logSheet.getLastRow() > 1) {
    const logData = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < logData.length; i++) {
      if (logData[i][0] === oldName) {
        logSheet.getRange(i + 2, 2).setValue(newName);
      }
    }
  }

  // 7. 備考ログの氏名を更新
  const notesSheet = ss.getSheetByName(NOTES_LOG_SHEET_NAME);
  if (notesSheet && notesSheet.getLastRow() > 1) {
    const notesData = notesSheet.getRange(2, 2, notesSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < notesData.length; i++) {
      if (notesData[i][0] === oldName) {
        notesSheet.getRange(i + 2, 2).setValue(newName);
      }
    }
  }

  return { success: true, data: { oldName: oldName, newName: newName, message: '「' + oldName + '」を「' + newName + '」に変更しました。' } };
}

function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// =====================================================================
//  曜日の数式を修復（TEXT→CHOOSE方式に変換）
// =====================================================================
function fixWeekdayFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '曜日の数式を修復',
    '全スタッフシートのB列（曜日）の数式を修復します。\n続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  ss.getSheets().forEach(sheet => {
    if (SYSTEM_SHEET_NAMES.includes(sheet.getName())) return;
    for (let i = 0; i < 31; i++) {
      const row = i + 5;
      sheet.getRange(row, 2).setFormula(
        '=IF(A' + row + '="","",CHOOSE(WEEKDAY(A' + row + '),"日","月","火","水","木","金","土"))'
      );
    }
  });

  ui.alert('完了', '全スタッフシートの曜日の数式を修復しました。', ui.ButtonSet.OK);
}

// =====================================================================
//  ヘルプ: スタッフ名変更
// =====================================================================
function showStaffHelp() {
  SpreadsheetApp.getUi().alert(
    'スタッフ名の変更方法',
    'スタッフ名を変更するには:\n\n' +
    '1. シートのタブ名を右クリック →「名前を変更」\n' +
    '2. シート内のA1セルのタイトルも合わせて修正\n\n' +
    'スタッフを追加するには:\n' +
    '1. 既存のスタッフシートのタブを右クリック →「コピーを作成」\n' +
    '2. タブ名とA1セルを新しいスタッフ名に変更\n' +
    '3. I列（備考）をクリア\n' +
    '4. 「スタッフ設定」シートに定時を設定',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =====================================================================
//  管理者API: 打刻ログ検索（指定日・スタッフ）
// =====================================================================
function apiGetClockLog_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return { success: false, error: '打刻ログシートが見つかりません' };

  const staffName = params.staffName || '';
  const dateParts = String(params.date).split('-');
  const targetDate = new Date(Number(dateParts[0]), Number(dateParts[1]) - 1, Number(dateParts[2]));
  targetDate.setHours(0, 0, 0, 0);

  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return { success: true, data: [] };

  const data = logSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const results = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const rowDate = new Date(row[3]);
    rowDate.setHours(0, 0, 0, 0);
    if (rowDate.getTime() !== targetDate.getTime()) continue;
    if (staffName && row[1] !== staffName) continue;

    results.push({
      rowIndex: i + 2,  // シート上の行番号（1-based、ヘッダー除く）
      timestamp: Utilities.formatDate(new Date(row[0]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
      staffName: row[1],
      type: row[2],
      date: Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy/MM/dd')
    });
  }

  return { success: true, data: results };
}

// =====================================================================
//  管理者API: 打刻追加（押し忘れ・有給）
// =====================================================================
function apiAddClockEntry_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return { success: false, error: '打刻ログシートが見つかりません' };

  const staffName = params.staffName;
  const type = params.type;  // '入室', '退室', '有給'
  if (!staffName || !type) return { success: false, error: 'スタッフ名と種別は必須です' };
  if (!['入室', '退室', '有給'].includes(type)) return { success: false, error: '種別が不正です: ' + type };

  const dateParts = String(params.date).split('-');
  const targetDate = new Date(Number(dateParts[0]), Number(dateParts[1]) - 1, Number(dateParts[2]));
  targetDate.setHours(0, 0, 0, 0);

  // 重複チェック
  if (logSheet.getLastRow() > 1) {
    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();
    for (const row of data) {
      if (!row[0]) continue;
      const rowDate = new Date(row[3]);
      rowDate.setHours(0, 0, 0, 0);
      if (rowDate.getTime() === targetDate.getTime() && row[1] === staffName && row[2] === type) {
        return { success: false, error: staffName + 'さんの同日・同種別の打刻が既に存在します' };
      }
    }
  }

  // 記録日時を組み立て
  let recordTime;
  if (type === '有給') {
    // 有給は対象日の0:00として記録
    recordTime = new Date(targetDate);
  } else {
    // 入室/退室は指定時刻で記録
    if (!params.time) return { success: false, error: '入室/退室の場合、時刻は必須です' };
    const timeParts = String(params.time).split(':');
    recordTime = new Date(targetDate);
    recordTime.setHours(Number(timeParts[0]), Number(timeParts[1]), 0, 0);
  }

  logSheet.appendRow([recordTime, staffName, type, targetDate]);
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow, 1).setNumberFormat('yyyy/MM/dd HH:mm');
  logSheet.getRange(lastRow, 4).setNumberFormat('yyyy/MM/dd');

  // 有給の場合、備考列にも記入
  if (type === '有給') {
    const staffSheet = ss.getSheetByName(staffName);
    if (staffSheet) {
      for (let i = 0; i < 31; i++) {
        const cellDate = staffSheet.getRange(i + 5, 1).getValue();
        if (cellDate instanceof Date) {
          const d = new Date(cellDate);
          d.setHours(0, 0, 0, 0);
          if (d.getTime() === targetDate.getTime()) {
            staffSheet.getRange(i + 5, 9).setValue('有給');
            break;
          }
        }
      }
    }
  }

  return { success: true, message: staffName + 'さんの' + type + 'を追加しました' };
}

// =====================================================================
//  管理者API: 打刻修正（時刻・種別の変更）
// =====================================================================
function apiEditClockEntry_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return { success: false, error: '打刻ログシートが見つかりません' };

  const rowIndex = Number(params.rowIndex);
  if (!rowIndex || rowIndex < 2) return { success: false, error: '行番号が不正です' };
  if (rowIndex > logSheet.getLastRow()) return { success: false, error: '指定行が存在しません' };

  // 楽観的排他制御: 現在値を検証
  const currentRow = logSheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
  const currentStaff = currentRow[1];
  const currentDate = new Date(currentRow[3]);
  currentDate.setHours(0, 0, 0, 0);

  if (params.expectedStaff && currentStaff !== params.expectedStaff) {
    return { success: false, error: 'データが他の操作で変更されています。再検索してください。' };
  }
  if (params.expectedDate) {
    const expParts = String(params.expectedDate).split('-');
    const expDate = new Date(Number(expParts[0]), Number(expParts[1]) - 1, Number(expParts[2]));
    expDate.setHours(0, 0, 0, 0);
    if (currentDate.getTime() !== expDate.getTime()) {
      return { success: false, error: 'データが他の操作で変更されています。再検索してください。' };
    }
  }

  // 新しい種別
  const newType = params.newType || currentRow[2];
  if (!['入室', '退室', '有給'].includes(newType)) return { success: false, error: '種別が不正です: ' + newType };

  // 新しい記録日時を組み立て
  let newRecordTime;
  if (newType === '有給') {
    newRecordTime = new Date(currentDate);
  } else {
    if (!params.newTime) return { success: false, error: '入室/退室の場合、時刻は必須です' };
    const timeParts = String(params.newTime).split(':');
    newRecordTime = new Date(currentDate);
    newRecordTime.setHours(Number(timeParts[0]), Number(timeParts[1]), 0, 0);
  }

  // 更新
  logSheet.getRange(rowIndex, 1).setValue(newRecordTime);
  logSheet.getRange(rowIndex, 1).setNumberFormat('yyyy/MM/dd HH:mm');
  logSheet.getRange(rowIndex, 3).setValue(newType);

  return { success: true, message: currentStaff + 'さんの打刻を修正しました' };
}

// =====================================================================
//  管理者API: 打刻削除
// =====================================================================
function apiDeleteClockEntry_(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return { success: false, error: '打刻ログシートが見つかりません' };

  const rowIndex = Number(params.rowIndex);
  if (!rowIndex || rowIndex < 2) return { success: false, error: '行番号が不正です' };
  if (rowIndex > logSheet.getLastRow()) return { success: false, error: '指定行が存在しません' };

  // 楽観的排他制御: 現在値を検証
  const currentRow = logSheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
  const currentStaff = currentRow[1];
  const currentType = currentRow[2];
  const currentDate = new Date(currentRow[3]);
  currentDate.setHours(0, 0, 0, 0);

  if (params.expectedStaff && currentStaff !== params.expectedStaff) {
    return { success: false, error: 'データが他の操作で変更されています。再検索してください。' };
  }
  if (params.expectedDate) {
    const expParts = String(params.expectedDate).split('-');
    const expDate = new Date(Number(expParts[0]), Number(expParts[1]) - 1, Number(expParts[2]));
    expDate.setHours(0, 0, 0, 0);
    if (currentDate.getTime() !== expDate.getTime()) {
      return { success: false, error: 'データが他の操作で変更されています。再検索してください。' };
    }
  }

  // 有給削除の場合、備考列もクリア
  if (currentType === '有給') {
    const staffSheet = ss.getSheetByName(currentStaff);
    if (staffSheet) {
      for (let i = 0; i < 31; i++) {
        const cellDate = staffSheet.getRange(i + 5, 1).getValue();
        if (cellDate instanceof Date) {
          const d = new Date(cellDate);
          d.setHours(0, 0, 0, 0);
          if (d.getTime() === currentDate.getTime()) {
            const note = staffSheet.getRange(i + 5, 9).getValue();
            if (String(note).indexOf('有給') !== -1) {
              staffSheet.getRange(i + 5, 9).clearContent();
            }
            break;
          }
        }
      }
    }
  }

  logSheet.deleteRow(rowIndex);
  return { success: true, message: currentStaff + 'さんの' + currentType + 'を削除しました' };
}

