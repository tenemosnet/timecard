/**
 * ====================================
 *  勤怠管理システム — Google Apps Script
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
const PDF_ROOT_FOLDER_NAME = '勤怠管理PDF';
const SYSTEM_SHEET_NAMES = [LOG_SHEET_NAME, NOTES_LOG_SHEET_NAME, SETTINGS_SHEET_NAME];
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
    .addItem('新しい月へ移行', 'promptNewMonth')
    .addItem('PDF出力（月次）', 'exportPdf')
    .addItem('備考を復元', 'restoreNotes')
    .addSeparator()
    .addItem('曜日の数式を修復', 'fixWeekdayFormulas')
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

  // スタッフ別シート作成（当月）
  const now = new Date();
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
  sheet.getRange('A1:B1').setValues([['スタッフ名', '定時(時間)']]);
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#e2e8f0');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 100);
  sheet.setFrozenRows(1);
  return sheet;
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
  sheet.getRange('C2').setValue('年：');
  sheet.getRange('D2').setValue(year).setFontWeight('bold');
  sheet.getRange('E2').setValue('月：');
  sheet.getRange('F2').setValue(month).setFontWeight('bold');
  sheet.getRange('H2').setValue('定時：');
  sheet.getRange('I2').setFormula(
    `=IFERROR(INDEX(FILTER('${SETTINGS_SHEET_NAME}'!B:B,'${SETTINGS_SHEET_NAME}'!A:A="${staffName}"),1),${DEFAULT_CONTRACTED_HOURS})`
  );

  // --- 行3: カラムヘッダー ---
  const headers = ['日', '曜日', '開始時間', '終了時間', '休憩時間', '定時内時間', '残業時間', '深夜残業時間', '備考'];
  sheet.getRange('A3:I3').setValues([headers]);
  sheet.getRange('A3:I3').setFontWeight('bold').setBackground('#e2e8f0')
    .setHorizontalAlignment('center');

  // --- 行4: 上部合計行 ---
  sheet.getRange('E4').setValue('合計').setFontWeight('bold').setHorizontalAlignment('center');
  // 合計の数式は下部合計を参照（行36）
  sheet.getRange('F4').setFormula('=F36').setNumberFormat('[h]:mm');
  // 上部合計行の注記なし（1分単位合計）
  sheet.getRange('A4:I4').setBackground('#f8f9fa');

  // --- 行5〜35: 日別データ（31行分、16日〜翌15日） ---
  for (let i = 0; i < 31; i++) {
    const row = i + 5;
    const dayOffset = i; // 0=16日目, 1=17日目, ...

    // A列: 日付
    // 前月16日からの日付を計算: DATE(year, month-1, 16+dayOffset)
    // 月の最終日（翌月15日）を超えたら空白
    sheet.getRange(row, 1).setFormula(
      `=LET(startDate,DATE(D2,F2-1,16),d,startDate+${dayOffset},endDate,DATE(D2,F2,${CUTOFF_DAY}),IF(d<=endDate,d,""))`
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

    // E列: 休憩時間（出退勤があれば固定1時間）
    sheet.getRange(row, 5).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),TIME(${BREAK_HOURS},0,0),"")`
    );
    sheet.getRange(row, 5).setNumberFormat('H:mm');

    // F列: 定時内時間 = MIN(実働時間, 定時)
    // 実働 = 終了 - 開始 - 休憩
    sheet.getRange(row, 6).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),MIN(D${row}-C${row}-E${row}, $I$2/24),"")`
    );
    sheet.getRange(row, 6).setNumberFormat('H:mm');

    // G列: 残業時間 = MAX(実働 - 定時, 0)
    sheet.getRange(row, 7).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),MAX(D${row}-C${row}-E${row}-$I$2/24, 0),"")`
    );
    sheet.getRange(row, 7).setNumberFormat('H:mm');

    // H列: 深夜残業時間（22:00〜5:00の勤務時間）
    // 終了時間が22:00を超えた場合: MIN(終了,翌5:00) - MAX(開始,22:00)
    sheet.getRange(row, 8).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),` +
      `IF(HOUR(D${row})>=22,(D${row}-TIME(22,0,0)),` +
      `IF(AND(HOUR(D${row})<5,D${row}<C${row}),(TIME(5,0,0)-TIME(0,0,0))-(TIME(22,0,0)-D${row}),` +
      `0)),"")`
    );
    sheet.getRange(row, 8).setNumberFormat('H:mm');

    // I列: 備考（手動入力用）
  }

  // --- 行36: 下部合計行 ---
  sheet.getRange('E36').setValue('合計').setFontWeight('bold').setHorizontalAlignment('center');
  // 1分単位の合計（定時内・残業・深夜残業）
  sheet.getRange('F36').setFormula('=SUMPRODUCT((F5:F35<>"")*F5:F35)');
  sheet.getRange('F36').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('G36').setFormula('=SUMPRODUCT((G5:G35<>"")*G5:G35)');
  sheet.getRange('G36').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('H36').setFormula('=SUMPRODUCT((H5:H35<>"")*H5:H35)');
  sheet.getRange('H36').setNumberFormat('[h]:mm').setFontWeight('bold');
  sheet.getRange('A36:I36').setBackground('#f8f9fa');

  // --- 列幅 ---
  [50, 40, 80, 80, 80, 90, 80, 90, 160].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // --- 条件付き書式: 土日の色分け ---
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

  sheet.setConditionalFormatRules([sundayRule, saturdayRule]);

  // ヘッダー固定
  sheet.setFrozenRows(4);

  return sheet;
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
//  新しい月へ移行
// =====================================================================
function promptNewMonth() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const yearRes = ui.prompt('新しい月へ移行', '年を入力してください（例: 2026）', ui.ButtonSet.OK_CANCEL);
  if (yearRes.getSelectedButton() !== ui.Button.OK) return;
  const year = parseInt(yearRes.getResponseText());

  const monthRes = ui.prompt('新しい月へ移行', '月を入力してください（例: 5）', ui.ButtonSet.OK_CANCEL);
  if (monthRes.getSelectedButton() !== ui.Button.OK) return;
  const month = parseInt(monthRes.getResponseText());

  if (isNaN(year) || isNaN(month) || month < 1 || month > 12) {
    ui.alert('エラー', '有効な年・月を入力してください。', ui.ButtonSet.OK);
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

    const curYear = sheet.getRange('D2').getValue();
    const curMonth = sheet.getRange('F2').getValue();

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
    sheet.getRange('D2').setValue(year);
    sheet.getRange('F2').setValue(month);
    sheet.getRange('I5:I35').clearContent();
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

  const year = staffSheets[0].getRange('D2').getValue();
  const month = staffSheets[0].getRange('F2').getValue();

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
  const fileName = staffName + '_' + year + '年' + monthStr + '月_勤怠表.pdf';

  // シートのD2/F2が指定年月と異なる場合、一時変更
  const origYear = sheet.getRange('D2').getValue();
  const origMonth = sheet.getRange('F2').getValue();
  const needRestore = (origYear != year || origMonth != month);

  if (needRestore) {
    sheet.getRange('D2').setValue(year);
    sheet.getRange('F2').setValue(month);
    SpreadsheetApp.flush();
  }

  // 同名ファイルが既にあれば削除
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
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
    '&range=A1:I36';

  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
  });

  const blob = response.getBlob().setName(fileName);
  const file = folder.createFile(blob);

  if (needRestore) {
    sheet.getRange('D2').setValue(origYear);
    sheet.getRange('F2').setValue(origMonth);
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

  const year = staffSheets[0].getRange('D2').getValue();
  const month = staffSheets[0].getRange('F2').getValue();
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
  const monthFolders = yearFolder.getFolders();

  while (monthFolders.hasNext()) {
    const mFolder = monthFolders.next();
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

// --- API: スタッフ一覧取得（定時設定付き） ---
function apiGetStaffList_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffNames = ss.getSheets()
    .filter(s => !s.isSheetHidden())
    .map(s => s.getName())
    .filter(n => !SYSTEM_SHEET_NAMES.includes(n));

  const staffList = staffNames.map(name => ({
    name: name,
    contractedHours: getStaffSetting_(ss, name)
  }));

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
    year = existingSheets[0].getRange('D2').getValue();
    month = existingSheets[0].getRange('F2').getValue();
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
