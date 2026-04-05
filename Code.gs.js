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
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // ← スプレッドシートのURLから取得して貼り付け
const STAFF_NAMES = ['山田 太郎', '佐藤 花子', '田中 一郎']; // ※ 初期設定時のみ使用。以降はシート名で管理
const LOG_SHEET_NAME = '打刻ログ';
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

    // B列: 曜日
    sheet.getRange(row, 2).setFormula(`=IF(A${row}="","",TEXT(A${row},"aaa"))`);

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

  // スタッフ名一覧はシート名から取得（打刻ログ以外）
  const sheets = ss.getSheets();
  const staffNames = sheets.map(s => s.getName()).filter(n => n !== LOG_SHEET_NAME);

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
    '備考欄の手入力データはクリアされます。\n\n' +
    '※ 打刻ログの生データは保持されます。\n' +
    '※ 前月のデータが必要な場合は先にCSVダウンロードしてください。\n\n' +
    '続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // 各スタッフシートの年月を更新
  ss.getSheets().forEach(sheet => {
    if (sheet.getName() === LOG_SHEET_NAME) return;
    sheet.getRange('D1').setValue(year);
    sheet.getRange('F1').setValue(month);
    sheet.getRange('H4:H34').clearContent(); // 備考欄クリア
  });

  ui.alert('完了', '全スタッフのシートを ' + year + '年' + month + '月に更新しました。', ui.ButtonSet.OK);
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
