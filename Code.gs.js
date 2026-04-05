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
 */

// =====================================================================
//  設定（ここを書き換えてカスタマイズ）
// =====================================================================
const SPREADSHEET_ID = '16ssw5Jfu3y8CepLMVMk_rnI4zBJBFUg8pnlfPYvFEPY';
const STAFF_NAMES = ['山田 太郎', '佐藤 花子', '田中 一郎']; // ※ 初期設定時のみ使用。以降はシート名で管理
const LOG_SHEET_NAME = '打刻ログ';
const NOTES_LOG_SHEET_NAME = '備考ログ';
const OVERTIME_THRESHOLD = 8;       // 残業基準（時間）
const LUNCH_DEDUCT_6H = 0.75;      // 6時間以上勤務: 45分控除
const LUNCH_DEDUCT_8H = 1.0;       // 8時間以上勤務: 60分控除

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

  // スタッフ別シート作成（当月）
  const now = new Date();
  STAFF_NAMES.forEach(name => {
    createStaffSheet_(ss, name, now.getFullYear(), now.getMonth() + 1);
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
    return sheet; // 既存なら何もしない
  }

  // ヘッダー
  sheet.getRange('A1:D1').setValues([['記録日時', '氏名', '種別', '日付']]);
  sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#e2e8f0');

  // 列幅
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 110);

  // ヘッダー固定
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

  // ヘッダー
  sheet.getRange('A1:E1').setValues([['年', '月', 'スタッフ名', '日', '備考内容']]);
  sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e2e8f0');

  // 列幅
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 40);
  sheet.setColumnWidth(5, 300);

  sheet.setFrozenRows(1);
  return sheet;
}

// =====================================================================
//  スタッフ別シート作成
// =====================================================================
function createStaffSheet_(ss, staffName, year, month) {
  // 既存シートがあればスキップ
  if (ss.getSheetByName(staffName)) return;

  const sheet = ss.insertSheet(staffName);
  const log = LOG_SHEET_NAME;

  // --- 行1: タイトル行 ---
  sheet.getRange('A1').setValue(staffName).setFontWeight('bold').setFontSize(13);
  sheet.getRange('C1').setValue('年：');
  sheet.getRange('D1').setValue(year).setFontWeight('bold');
  sheet.getRange('E1').setValue('月：');
  sheet.getRange('F1').setValue(month).setFontWeight('bold');

  // --- 行3: ヘッダー ---
  const headers = ['日付', '曜日', '入室', '退室', '実働(h)', '通常(h)', '残業(h)', '備考'];
  sheet.getRange('A3:H3').setValues([headers]);
  sheet.getRange('A3:H3').setFontWeight('bold').setBackground('#e2e8f0');

  // --- 行4〜34: 日別データ（最大31日分） ---
  for (let i = 0; i < 31; i++) {
    const row = i + 4;
    const dayNum = i + 1;

    // A列: 日付（月の日数を超えたら空白）
    sheet.getRange(row, 1).setFormula(
      `=IF(${dayNum}<=DAY(EOMONTH(DATE(D1,F1,1),0)),DATE(D1,F1,${dayNum}),"")`
    );
    sheet.getRange(row, 1).setNumberFormat('m/d');

    // B列: 曜日（地域設定に依存しない方式）
    sheet.getRange(row, 2).setFormula(`=IF(A${row}="","",CHOOSE(WEEKDAY(A${row}),"日","月","火","水","木","金","土"))`);

    // C列: 入室時刻（打刻ログから自動取得）
    sheet.getRange(row, 3).setFormula(
      `=IFERROR(INDEX(FILTER('${log}'!$A:$A,'${log}'!$D:$D=A${row},'${log}'!$B:$B=$A$1,'${log}'!$C:$C="入室"),1),"")`
    );
    sheet.getRange(row, 3).setNumberFormat('H:mm');

    // D列: 退室時刻（打刻ログから自動取得）
    sheet.getRange(row, 4).setFormula(
      `=IFERROR(INDEX(FILTER('${log}'!$A:$A,'${log}'!$D:$D=A${row},'${log}'!$B:$B=$A$1,'${log}'!$C:$C="退室"),1),"")`
    );
    sheet.getRange(row, 4).setNumberFormat('H:mm');

    // E列: 実働時間（昼休み控除後）
    sheet.getRange(row, 5).setFormula(
      `=IF(AND(C${row}<>"",D${row}<>""),LET(raw,(D${row}-C${row})*24,lunch,IF(raw>=${OVERTIME_THRESHOLD},${LUNCH_DEDUCT_8H},IF(raw>=6,${LUNCH_DEDUCT_6H},0)),MAX(raw-lunch,0)),"")`
    );
    sheet.getRange(row, 5).setNumberFormat('0.00');

    // F列: 通常勤務
    sheet.getRange(row, 6).setFormula(`=IF(E${row}="","",MIN(E${row},${OVERTIME_THRESHOLD}))`);
    sheet.getRange(row, 6).setNumberFormat('0.00');

    // G列: 残業
    sheet.getRange(row, 7).setFormula(`=IF(E${row}="","",MAX(E${row}-${OVERTIME_THRESHOLD},0))`);
    sheet.getRange(row, 7).setNumberFormat('0.00');

    // H列: 備考（手動入力用、空白のまま）
  }

  // --- 行36〜41: 月次集計 ---
  sheet.getRange('A36').setValue('【月次集計】').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A36:H36').setBackground('#f1f5f9');

  const summaryRows = [
    ['出勤日数',       '=COUNTA(FILTER(C4:C34,C4:C34<>""))',  '日'],
    ['通常勤務合計',   '=SUM(F4:F34)',                          '時間'],
    ['残業合計',       '=SUM(G4:G34)',                          '時間'],
    ['昼休み控除合計', `=SUMPRODUCT((E4:E34<>"")*(IF((D4:D34-C4:C34)*24>=${OVERTIME_THRESHOLD},${LUNCH_DEDUCT_8H},IF((D4:D34-C4:C34)*24>=6,${LUNCH_DEDUCT_6H},0))))`, '時間'],
    ['総勤務時間',     '=B38+B39',                              '時間'],
  ];

  summaryRows.forEach((r, i) => {
    const row = 37 + i;
    sheet.getRange(row, 1).setValue(r[0]).setFontWeight('bold');
    sheet.getRange(row, 2).setFormula(r[1]).setNumberFormat('0.00');
    sheet.getRange(row, 3).setValue(r[2]);
  });

  // 総勤務時間を強調
  sheet.getRange('A41:C41').setFontWeight('bold').setBackground('#dbeafe');

  // --- 列幅 ---
  [80, 50, 80, 80, 80, 80, 80, 160].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // --- 条件付き書式: 土日の色分け ---
  const dataRange = sheet.getRange('A4:H34');

  const sundayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=WEEKDAY($A4)=1')
    .setFontColor('#dc2626')
    .setBackground('#fef2f2')
    .setRanges([dataRange])
    .build();

  const saturdayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=WEEKDAY($A4)=7')
    .setFontColor('#2563eb')
    .setBackground('#eff6ff')
    .setRanges([dataRange])
    .build();

  sheet.setConditionalFormatRules([sundayRule, saturdayRule]);

  // ヘッダー固定
  sheet.setFrozenRows(3);

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
      const rowDate = new Date(row[3]); // D列: 日付
      if (rowDate.getTime() === today.getTime() && row[1] === staffName && row[2] === type) {
        throw new Error(staffName + 'さんは本日すでに' + type + '打刻済みです');
      }
    }
  }

  // 打刻ログに追記
  logSheet.appendRow([now, staffName, type, today]);

  // セルの書式設定
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

  // 打刻ログから本日分を取得
  let todayLogs = [];
  if (logSheet && logSheet.getLastRow() > 1) {
    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();
    todayLogs = data.filter(row => {
      if (!row[3]) return false;
      const d = new Date(row[3]);
      return d.getTime() === today.getTime();
    });
  }

  // スタッフ名一覧はシート名から取得（システムシートを除外）
  const sheets = ss.getSheets();
  const staffNames = sheets.map(s => s.getName()).filter(n => n !== LOG_SHEET_NAME && n !== NOTES_LOG_SHEET_NAME);

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

  // 確認
  const confirm = ui.alert(
    '確認',
    '各スタッフのシートを ' + year + '年' + month + '月に切り替えます。\n' +
    '備考欄は自動バックアップ後にクリアされます。\n\n' +
    '※ 打刻ログの生データは保持されます。\n' +
    '※ 備考は「備考ログ」シートに保存され、後から復元できます。\n\n' +
    '続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // 備考ログシートを取得（なければ作成）
  let notesLog = ss.getSheetByName(NOTES_LOG_SHEET_NAME);
  if (!notesLog) notesLog = createNotesLogSheet_(ss);

  // 各スタッフシートの備考をバックアップ → 年月更新
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === LOG_SHEET_NAME || name === NOTES_LOG_SHEET_NAME) return;

    // 現在の年月を取得（移行前）
    const curYear = sheet.getRange('D1').getValue();
    const curMonth = sheet.getRange('F1').getValue();

    // 備考欄をバックアップ
    const notes = sheet.getRange('H4:H34').getValues();
    const rows = [];
    notes.forEach((cell, i) => {
      if (cell[0] !== '' && cell[0] !== null) {
        rows.push([curYear, curMonth, name, i + 1, cell[0]]);
      }
    });
    if (rows.length > 0) {
      // 同じ年月・スタッフの既存バックアップを削除（上書き相当）
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
    sheet.getRange('D1').setValue(year);
    sheet.getRange('F1').setValue(month);
    sheet.getRange('H4:H34').clearContent();
  });

  ui.alert('完了', '全スタッフのシートを ' + year + '年' + month + '月に更新しました。', ui.ButtonSet.OK);
}

// =====================================================================
//  PDF出力（月次）
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
    s.getName() !== LOG_SHEET_NAME && s.getName() !== NOTES_LOG_SHEET_NAME
  );
  if (staffSheets.length === 0) return;

  // 最初のスタッフシートから現在の表示年月を取得
  const year = staffSheets[0].getRange('D1').getValue();
  const month = staffSheets[0].getRange('F1').getValue();

  const confirm = ui.alert(
    '備考を復元',
    year + '年' + month + '月の備考データを復元します。\n' +
    '現在の備考欄は上書きされます。続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // 備考ログから該当年月のデータを読み取り
  const logData = notesLog.getRange(2, 1, notesLog.getLastRow() - 1, 5).getValues();
  let restoredCount = 0;

  staffSheets.forEach(sheet => {
    const staffName = sheet.getName();
    const matches = logData.filter(row =>
      row[0] == year && row[1] == month && row[2] === staffName
    );
    matches.forEach(row => {
      const day = row[3]; // 日（1〜31）
      const note = row[4];
      if (day >= 1 && day <= 31) {
        sheet.getRange(day + 3, 8).setValue(note); // H列、行4が1日
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
//  PDF出力（月次）
// =====================================================================
function exportPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // スタッフシート一覧を取得
  const staffSheets = ss.getSheets().filter(s => s.getName() !== LOG_SHEET_NAME && s.getName() !== NOTES_LOG_SHEET_NAME);
  if (staffSheets.length === 0) {
    ui.alert('エラー', 'スタッフシートが見つかりません。先に初期設定を実行してください。', ui.ButtonSet.OK);
    return;
  }

  // 最初のスタッフシートから年月を取得
  const year = staffSheets[0].getRange('D1').getValue();
  const month = staffSheets[0].getRange('F1').getValue();
  const staffNames = staffSheets.map(s => s.getName()).join('、');

  const confirm = ui.alert(
    'PDF出力',
    year + '年' + month + '月の勤務表をPDFで出力します。\n\n' +
    '対象スタッフ: ' + staffNames + '\n\n' +
    'Googleドライブに保存されます。続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // 出力先フォルダを作成（タイムカードPDF / YYYY年MM月）
  const rootFolderName = 'タイムカードPDF';
  const monthFolderName = year + '年' + String(month).padStart(2, '0') + '月';

  let rootFolder;
  const rootFolders = DriveApp.getFoldersByName(rootFolderName);
  if (rootFolders.hasNext()) {
    rootFolder = rootFolders.next();
  } else {
    rootFolder = DriveApp.createFolder(rootFolderName);
  }

  let monthFolder;
  const monthFolders = rootFolder.getFoldersByName(monthFolderName);
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = rootFolder.createFolder(monthFolderName);
  }

  // 各スタッフシートをPDF出力
  const ssId = ss.getId();
  const token = ScriptApp.getOAuthToken();
  const createdFiles = [];

  staffSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const fileName = sheetName + '_' + year + '年' + String(month).padStart(2, '0') + '月_勤務表.pdf';

    // 同名ファイルが既にあれば削除（上書き相当）
    const existing = monthFolder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    // PDF生成URL
    const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' +
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
      '&range=A1:H41';

    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });

    const blob = response.getBlob().setName(fileName);
    const file = monthFolder.createFile(blob);
    createdFiles.push(sheetName);
  });

  ui.alert(
    'PDF出力完了',
    createdFiles.length + '名分のPDFを出力しました。\n\n' +
    '保存先: Googleドライブ\n' +
    '  📁 ' + rootFolderName + ' / ' + monthFolderName + '\n\n' +
    '出力ファイル:\n' +
    createdFiles.map(name => '  ・' + name).join('\n'),
    ui.ButtonSet.OK
  );
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
    if (sheet.getName() === LOG_SHEET_NAME) return;
    for (let i = 0; i < 31; i++) {
      const row = i + 4;
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
    '2. シート内のA1セルも同じ名前に変更\n\n' +
    'スタッフを追加するには:\n' +
    '1. 既存のスタッフシートのタブを右クリック →「コピーを作成」\n' +
    '2. タブ名とA1セルを新しいスタッフ名に変更\n' +
    '3. H列（備考）をクリア',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
