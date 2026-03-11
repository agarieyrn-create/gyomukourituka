// ============================================================
// setup.gs - セットアップ・カスタムメニュー
// GAS 予約管理ツール
// ============================================================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ 予約ツール設定')
    .addItem('初回セットアップ', 'setupTool')
    .addSeparator()
    .addItem('ダッシュボードを今すぐ更新', 'updateDashboard')
    .addItem('リマインダートリガーを設定', 'setupReminderTrigger')
    .addItem('臨時休業日を追加', 'addTemporaryHoliday')
    .addSeparator()
    .addItem('テストメール送信', 'testSendMail')
    .addToUi();
}

/**
 * 初回セットアップ
 * 必要なシートとヘッダーを自動作成する
 */
function setupTool() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '初回セットアップ',
    '予約管理ツールの初回セットアップを実行します。\n' +
    '必要なシートとヘッダーが自動作成されます。\n\n' +
    '既存のシートがある場合は上書きされません。\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 予約一覧シート ---
  var reservationSheet = getOrCreateSheet_(ss, '予約一覧');
  if (reservationSheet.getLastRow() === 0) {
    reservationSheet.appendRow([
      '予約ID', '受付日時', '予約日', '予約時間', '氏名', '電話番号',
      'メールアドレス', 'メニュー', '料金', 'ステータス', '備考',
      'カレンダーイベントID', 'リマインダー送信済みフラグ'
    ]);
    formatHeader_(reservationSheet);
  }

  // --- メニューマスタシート ---
  var menuSheet = getOrCreateSheet_(ss, 'メニューマスタ');
  if (menuSheet.getLastRow() === 0) {
    menuSheet.appendRow(['メニューID', 'メニュー名', '所要時間（分）', '料金']);
    formatHeader_(menuSheet);

    // サンプルデータ
    menuSheet.appendRow(['MENU-001', 'カット', 60, 4000]);
    menuSheet.appendRow(['MENU-002', 'カット＋カラー', 120, 8000]);
    menuSheet.appendRow(['MENU-003', 'パーマ', 90, 7000]);
    menuSheet.appendRow(['MENU-004', 'トリートメント', 30, 3000]);
    menuSheet.appendRow(['MENU-005', 'ヘッドスパ', 60, 5000]);
  }

  // --- 設定シート ---
  var settingsSheet = getOrCreateSheet_(ss, '設定');
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.appendRow(['設定項目', '設定値']);
    formatHeader_(settingsSheet);

    // デフォルト設定値
    var defaultSettings = [
      ['営業開始時間', '10:00'],
      ['営業終了時間', '19:00'],
      ['予約間隔（分）', 60],
      ['同時受付数', 1],
      ['定休日曜日', '月曜,火曜'],
      ['管理者メールアドレス', ''],
      ['店舗名', 'サンプルサロン'],
      ['共通GoogleカレンダーID', ''],
      ['予約受付期間（何日先まで）', 30],
      ['リマインダー送信タイミング', '1日前']
    ];

    for (var s = 0; s < defaultSettings.length; s++) {
      settingsSheet.appendRow(defaultSettings[s]);
    }

    // 列幅調整
    settingsSheet.setColumnWidth(1, 250);
    settingsSheet.setColumnWidth(2, 300);
  }

  // --- 臨時休業日シート ---
  var holidaySheet = getOrCreateSheet_(ss, '臨時休業日');
  if (holidaySheet.getLastRow() === 0) {
    holidaySheet.appendRow(['日付', '理由']);
    formatHeader_(holidaySheet);

    // サンプルデータ
    holidaySheet.appendRow([new Date(2026, 3, 29), '設備メンテナンス']);
    holidaySheet.appendRow([new Date(2026, 4, 3), '研修のため']);
    holidaySheet.appendRow([new Date(2026, 11, 30), '年末休暇']);
    holidaySheet.appendRow([new Date(2026, 11, 31), '年末休暇']);
  }

  // --- キャンセル・変更履歴シート ---
  var historySheet = getOrCreateSheet_(ss, 'キャンセル・変更履歴');
  if (historySheet.getLastRow() === 0) {
    historySheet.appendRow([
      '処理日時', '予約ID', '氏名', '元の予約日時', '処理種別', '変更後日時'
    ]);
    formatHeader_(historySheet);
  }

  // --- ダッシュボードシート ---
  getOrCreateSheet_(ss, 'ダッシュボード');

  // ダッシュボードを初回更新
  updateDashboard();

  // リマインダートリガーを設定
  setupReminderTrigger();

  ui.alert(
    'セットアップ完了',
    '予約管理ツールのセットアップが完了しました。\n\n' +
    '次のステップ：\n' +
    '1. 「設定」シートで店舗情報を入力してください\n' +
    '2. 「メニューマスタ」シートでメニューを編集してください\n' +
    '3. Web Appとしてデプロイしてください\n\n' +
    'リマインダートリガー（毎日8:00）も自動設定されました。',
    ui.ButtonSet.OK
  );
}

/**
 * テストメール送信
 * 設定シートの管理者メールアドレスにテストメールを送信する
 */
function testSendMail() {
  var ui = SpreadsheetApp.getUi();
  var settings = getSettings();
  var adminEmail = settings['管理者メールアドレス'];

  if (!adminEmail) {
    ui.alert('エラー', '設定シートに管理者メールアドレスを入力してください。', ui.ButtonSet.OK);
    return;
  }

  var shopName = settings['店舗名'] || '当店';

  try {
    // テスト用の予約データ
    var testReservation = {
      reservationId: 'RSV-TEST-001',
      date: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
      time: '14:00',
      name: 'テスト太郎',
      phone: '090-1234-5678',
      email: adminEmail,
      menuName: 'テストメニュー',
      price: 5000
    };

    sendConfirmationMail(testReservation);

    ui.alert('送信完了', adminEmail + ' にテストメールを送信しました。\n受信を確認してください。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', 'メール送信に失敗しました: ' + e.message, ui.ButtonSet.OK);
  }
}

// ============================================================
// 内部ヘルパー関数
// ============================================================

/**
 * シートを取得、なければ新規作成する（内部関数）
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @return {Sheet} シートオブジェクト
 */
function getOrCreateSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * ヘッダー行の書式を設定する（内部関数）
 * @param {Sheet} sheet - シートオブジェクト
 */
function formatHeader_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol <= 0) return;

  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#4472C4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 行を固定
  sheet.setFrozenRows(1);
}
