// ============================================================
// setup.gs - 初回セットアップ・シート初期化
// 請求書・見積書 自動生成ツール
// ============================================================

/**
 * 初回セットアップ用の関数
 * 自社情報をダイアログで入力し、スクリプトプロパティに保存する
 * カスタムメニューの「初回セットアップ」から呼び出されます
 */
function setupProperties() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  // 現在の設定値を取得（既に設定済みの場合はデフォルト値として表示）
  var currentCompany = props.getProperty('COMPANY_NAME') || '';
  var currentZip = props.getProperty('ZIP_CODE') || '';
  var currentAddress = props.getProperty('ADDRESS') || '';
  var currentPhone = props.getProperty('PHONE') || '';
  var currentEmail = props.getProperty('EMAIL') || '';
  var currentBank = props.getProperty('BANK_INFO') || '';

  // --- 会社名 ---
  var companyResult = ui.prompt(
    '初回セットアップ（1/6）',
    '自社の会社名（または屋号・氏名）を入力してください。\n' +
    (currentCompany ? '現在の設定: ' + currentCompany : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (companyResult.getSelectedButton() !== ui.Button.OK) return;
  var companyName = companyResult.getResponseText() || currentCompany;

  // --- 郵便番号 ---
  var zipResult = ui.prompt(
    '初回セットアップ（2/6）',
    '郵便番号を入力してください。（例: 100-0001）\n' +
    (currentZip ? '現在の設定: ' + currentZip : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (zipResult.getSelectedButton() !== ui.Button.OK) return;
  var zipCode = zipResult.getResponseText() || currentZip;

  // --- 住所 ---
  var addressResult = ui.prompt(
    '初回セットアップ（3/6）',
    '住所を入力してください。\n' +
    (currentAddress ? '現在の設定: ' + currentAddress : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (addressResult.getSelectedButton() !== ui.Button.OK) return;
  var address = addressResult.getResponseText() || currentAddress;

  // --- 電話番号 ---
  var phoneResult = ui.prompt(
    '初回セットアップ（4/6）',
    '電話番号を入力してください。\n' +
    (currentPhone ? '現在の設定: ' + currentPhone : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (phoneResult.getSelectedButton() !== ui.Button.OK) return;
  var phone = phoneResult.getResponseText() || currentPhone;

  // --- メールアドレス ---
  var emailResult = ui.prompt(
    '初回セットアップ（5/6）',
    'メールアドレスを入力してください。\n' +
    (currentEmail ? '現在の設定: ' + currentEmail : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (emailResult.getSelectedButton() !== ui.Button.OK) return;
  var email = emailResult.getResponseText() || currentEmail;

  // --- 振込先情報 ---
  var bankResult = ui.prompt(
    '初回セットアップ（6/6）',
    '振込先情報を入力してください。\n' +
    '例: ○○銀行 △△支店 普通 1234567 口座名義\n' +
    (currentBank ? '現在の設定: ' + currentBank : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (bankResult.getSelectedButton() !== ui.Button.OK) return;
  var bankInfo = bankResult.getResponseText() || currentBank;

  // --- スクリプトプロパティに保存 ---
  props.setProperties({
    'COMPANY_NAME': companyName,
    'ZIP_CODE': zipCode,
    'ADDRESS': address,
    'PHONE': phone,
    'EMAIL': email,
    'BANK_INFO': bankInfo
  });

  ui.alert(
    'セットアップ完了',
    '自社情報を保存しました。\n\n' +
    '会社名: ' + companyName + '\n' +
    '郵便番号: ' + zipCode + '\n' +
    '住所: ' + address + '\n' +
    '電話番号: ' + phone + '\n' +
    'メール: ' + email + '\n' +
    '振込先: ' + bankInfo,
    ui.ButtonSet.OK
  );
}

/**
 * スプレッドシートの全シートを初期化する
 * ※この関数は手動で1回だけ実行してください
 * 顧客マスタ、品目マスタ、入力フォーム、出力シート、履歴管理の5シートを作成します
 */
function initializeSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    // --- 既存データの確認 ---
    var existingSheets = ['顧客マスタ', '品目マスタ'];
    var hasData = false;
    for (var s = 0; s < existingSheets.length; s++) {
      var existing = ss.getSheetByName(existingSheets[s]);
      if (existing && existing.getLastRow() > 1) {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      var confirm = ui.alert(
        '⚠️ データが上書きされます',
        '「顧客マスタ」「品目マスタ」などの既存データがすべて削除されます。\n\n' +
        '本当に初期化しますか？\n' +
        '（実行前にデータをバックアップしてください）',
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }

    // --- 1. 顧客マスタ ---
    var customerSheet = getOrCreateSheet_(ss, '顧客マスタ');
    setupCustomerMaster_(customerSheet);

    // --- 2. 品目マスタ ---
    var itemSheet = getOrCreateSheet_(ss, '品目マスタ');
    setupItemMaster_(itemSheet);

    // --- 3. 入力フォーム ---
    var formSheet = getOrCreateSheet_(ss, '入力フォーム');
    setupInputForm_(ss, formSheet);

    // --- 4. 出力シート ---
    getOrCreateSheet_(ss, '出力シート');

    // --- 5. 履歴管理 ---
    var historySheet = getOrCreateSheet_(ss, '履歴管理');
    setupHistorySheet_(historySheet);

    // --- デフォルトの「シート1」を削除（存在する場合） ---
    var defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }

    // 入力フォームをアクティブにする
    ss.setActiveSheet(ss.getSheetByName('入力フォーム'));

    ui.alert(
      '初期化完了',
      'スプレッドシートの初期化が完了しました。\n\n' +
      '次のステップ:\n' +
      '1. メニュー「📄 書類生成ツール」→「初回セットアップ」で自社情報を登録\n' +
      '2. 「顧客マスタ」「品目マスタ」にデータを追加\n' +
      '3. 「入力フォーム」から書類を作成',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('初期化エラー', e.message, ui.ButtonSet.OK);
    Logger.log('initializeSpreadsheet エラー: ' + e.message + '\n' + e.stack);
  }
}

// ============================================================
// シート初期化の内部関数
// ============================================================

/**
 * シートを取得する。存在しなければ新規作成する（内部関数）
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} sheetName - シート名
 * @return {Sheet} シートオブジェクト
 */
function getOrCreateSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    sheet.clearFormats();
  }
  return sheet;
}

/**
 * 顧客マスタシートを設定する（内部関数）
 * @param {Sheet} sheet - 顧客マスタシート
 */
function setupCustomerMaster_(sheet) {
  // ヘッダー
  var headers = ['顧客ID', '会社名', '担当者名', '郵便番号', '住所', 'メールアドレス', '電話番号'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4472C4');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');

  // サンプルデータ
  var sampleData = [
    ['C001', '株式会社サンプル商事', '山田太郎', '100-0001', '東京都千代田区千代田1-1-1', 'yamada@sample.co.jp', '03-1234-5678'],
    ['C002', '合同会社テスト工業', '鈴木花子', '530-0001', '大阪府大阪市北区梅田2-2-2', 'suzuki@test-ind.co.jp', '06-9876-5432'],
    ['C003', '有限会社デモサービス', '佐藤一郎', '460-0008', '愛知県名古屋市中区栄3-3-3', 'sato@demo-svc.co.jp', '052-1111-2222']
  ];
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  // 列幅調整
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 250);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 130);

  // 行を固定
  sheet.setFrozenRows(1);
}

/**
 * 品目マスタシートを設定する（内部関数）
 * @param {Sheet} sheet - 品目マスタシート
 */
function setupItemMaster_(sheet) {
  // ヘッダー
  var headers = ['品目ID', '品名', '単価', '単位', '税区分'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4472C4');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');

  // サンプルデータ
  var sampleData = [
    ['I001', 'Webサイト制作（基本プラン）', 300000, '式', '10%'],
    ['I002', 'Webサイト制作（プレミアムプラン）', 500000, '式', '10%'],
    ['I003', 'ロゴデザイン', 80000, '式', '10%'],
    ['I004', 'バナー制作', 15000, '個', '10%'],
    ['I005', 'コンサルティング', 10000, '時間', '10%'],
    ['I006', 'システム開発', 8000, '時間', '10%'],
    ['I007', 'サーバー保守・運用', 50000, '式', '10%'],
    ['I008', '記事執筆', 5000, '個', '10%'],
    ['I009', '翻訳（日→英）', 3000, '個', '10%'],
    ['I010', '食品加工費', 20000, '式', '8%']
  ];
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  // 列幅調整
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 80);

  // 単価フォーマット
  sheet.getRange(2, 3, sampleData.length, 1).setNumberFormat('#,##0');

  // 行を固定
  sheet.setFrozenRows(1);
}

/**
 * 入力フォームシートを設定する（内部関数）
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Sheet} sheet - 入力フォームシート
 */
function setupInputForm_(ss, sheet) {
  // 列幅設定
  sheet.setColumnWidth(1, 20);   // A: 余白
  sheet.setColumnWidth(2, 120);  // B: ラベル
  sheet.setColumnWidth(3, 200);  // C: 入力欄（品目名）
  sheet.setColumnWidth(4, 100);  // D: 単価（自動）
  sheet.setColumnWidth(5, 80);   // E: 数量
  sheet.setColumnWidth(6, 60);   // F: 単位（自動）
  sheet.setColumnWidth(7, 120);  // G: 金額

  // --- タイトル ---
  var titleRange = sheet.getRange('B1:G1');
  titleRange.merge();
  titleRange.setValue('請求書・見積書 入力フォーム');
  titleRange.setFontSize(16);
  titleRange.setFontWeight('bold');
  titleRange.setBackground('#4472C4');
  titleRange.setFontColor('#FFFFFF');
  titleRange.setHorizontalAlignment('center');

  // --- 操作ガイド（初心者向けステップ表示） ---
  var guideRange = sheet.getRange('B2:G2');
  guideRange.merge();
  guideRange.setValue(
    '【操作手順】 ① 書類種別を選択 → ② 顧客を選択 → ③ 品目をドロップダウンから選択・数量を入力 → ' +
    '④ 必要に応じて発行日・支払期限・備考を編集 → ⑤ メニュー「📄 書類生成ツール」→「書類を生成する」'
  );
  guideRange.setFontSize(9);
  guideRange.setFontColor('#333333');
  guideRange.setBackground('#E8F0FE');
  guideRange.setWrap(true);
  sheet.setRowHeight(2, 36);

  // --- 書類種別 ---
  sheet.getRange('B3').setValue('書類種別');
  sheet.getRange('B3').setFontWeight('bold');
  // ドロップダウンを設定
  var docTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['請求書', '見積書'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C3').setDataValidation(docTypeRule);
  sheet.getRange('C3').setBackground('#FFF2CC');

  // --- 顧客名 ---
  sheet.getRange('B5').setValue('顧客名');
  sheet.getRange('B5').setFontWeight('bold');
  // 顧客マスタの会社名からドロップダウンを作成（末尾行まで動的に取得）
  var customerMasterSheet = ss.getSheetByName('顧客マスタ');
  var customerLastRow = Math.max(customerMasterSheet.getLastRow(), 2);
  var customerRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(customerMasterSheet.getRange('B2:B' + customerLastRow), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C5').setDataValidation(customerRule);
  sheet.getRange('C5').setBackground('#FFF2CC');

  // --- 顧客情報プレビュー（選択した顧客の住所・担当者を自動表示） ---
  // 顧客を選択すると、その下に担当者と住所が自動で表示される
  sheet.getRange('D5:G5').merge();
  sheet.getRange('D5').setFormula(
    '=IFERROR("担当: "&VLOOKUP(C5,顧客マスタ!B:C,2,FALSE)&"　住所: "&VLOOKUP(C5,顧客マスタ!B:E,4,FALSE),"")'
  );
  sheet.getRange('D5').setFontSize(9);
  sheet.getRange('D5').setFontColor('#666666');

  // --- 発行日 ---
  sheet.getRange('B7').setValue('発行日');
  sheet.getRange('B7').setFontWeight('bold');
  sheet.getRange('C7').setNumberFormat('yyyy/MM/dd');
  sheet.getRange('C7').setBackground('#FFF2CC');
  // デフォルトで今日の日付をセット（初心者が入力忘れしにくいように）
  sheet.getRange('C7').setValue(new Date());

  // --- 支払期限（発行日+30日で自動入力。変更も可能） ---
  sheet.getRange('B9').setValue('支払期限');
  sheet.getRange('B9').setFontWeight('bold');
  sheet.getRange('C9').setFormula('=IF(C7="","",C7+30)');
  sheet.getRange('C9').setNumberFormat('yyyy/MM/dd');
  sheet.getRange('C9').setBackground('#E2EFDA');  // 自動計算（緑）
  sheet.getRange('D9:G9').merge();
  sheet.getRange('D9').setValue('※請求書のみ（発行日+30日で自動入力。直接入力で変更可）');
  sheet.getRange('D9').setFontColor('#888888');
  sheet.getRange('D9').setFontSize(9);

  // --- 品目セクション ---
  sheet.getRange('B11').setValue('品目明細');
  sheet.getRange('B11').setFontWeight('bold');
  sheet.getRange('B11').setFontSize(12);

  // 品目ヘッダー
  sheet.getRange('B12').setValue('No.');
  sheet.getRange('C12').setValue('品名');
  sheet.getRange('D12').setValue('単価');
  sheet.getRange('E12').setValue('数量');
  sheet.getRange('F12').setValue('単位');
  sheet.getRange('G12').setValue('金額');
  sheet.getRange('B12:G12').setFontWeight('bold');
  sheet.getRange('B12:G12').setBackground('#D9E2F3');

  // 品目マスタの品名からドロップダウンルールを作成（末尾行まで動的に取得）
  var itemMasterSheet = ss.getSheetByName('品目マスタ');
  var itemLastRow = Math.max(itemMasterSheet.getLastRow(), 2);
  var itemRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(itemMasterSheet.getRange('B2:B' + itemLastRow), true)
    .setAllowInvalid(false)
    .build();

  // 品目1〜10を配置
  for (var i = 0; i < 10; i++) {
    var row = 13 + (i * 2);

    // 番号
    sheet.getRange('B' + row).setValue(i + 1);
    sheet.getRange('B' + row).setHorizontalAlignment('center');

    // 品名（ドロップダウン）
    sheet.getRange('C' + row).setDataValidation(itemRule);
    sheet.getRange('C' + row).setBackground('#FFF2CC');

    // 単価（品目マスタから自動表示）
    var unitPriceFormula = '=IFERROR(VLOOKUP(C' + row + ',品目マスタ!B:C,2,FALSE),"")';
    sheet.getRange('D' + row).setFormula(unitPriceFormula);
    sheet.getRange('D' + row).setNumberFormat('#,##0');
    sheet.getRange('D' + row).setHorizontalAlignment('right');
    sheet.getRange('D' + row).setBackground('#E2EFDA');

    // 数量
    sheet.getRange('E' + row).setBackground('#FFF2CC');
    sheet.getRange('E' + row).setNumberFormat('#,##0');

    // 単位（品目マスタから自動表示）
    var unitFormula = '=IFERROR(VLOOKUP(C' + row + ',品目マスタ!B:D,3,FALSE),"")';
    sheet.getRange('F' + row).setFormula(unitFormula);
    sheet.getRange('F' + row).setHorizontalAlignment('center');
    sheet.getRange('F' + row).setBackground('#E2EFDA');

    // 金額（自動計算の数式を設定）
    // VLOOKUP で品目マスタから単価を取得し、数量と掛け算
    var amountFormula = '=IFERROR(VLOOKUP(C' + row + ',品目マスタ!B:C,2,FALSE)*E' + row + ',"")';
    sheet.getRange('G' + row).setFormula(amountFormula);
    sheet.getRange('G' + row).setNumberFormat('#,##0');
    sheet.getRange('G' + row).setBackground('#E2EFDA');

    // 偶数行に背景色（入力セルはやや濃い黄、自動計算セルはやや濃い緑）
    if (i % 2 === 1) {
      sheet.getRange('B' + row + ':G' + row).setBackground('#F2F2F2');
      sheet.getRange('C' + row).setBackground('#FFF8E1');
      sheet.getRange('D' + row).setBackground('#D5E8D4');
      sheet.getRange('E' + row).setBackground('#FFF8E1');
      sheet.getRange('F' + row).setBackground('#D5E8D4');
      sheet.getRange('G' + row).setBackground('#D5E8D4');
    }
  }

  // --- リアルタイム合計欄（入力中に金額を確認できる） ---
  // 品目の最終行（品目10 = row 31）の次の行
  sheet.getRange('E32').setValue('小計（税抜）:');
  sheet.getRange('E32').setFontWeight('bold');
  sheet.getRange('E32').setHorizontalAlignment('right');
  // 全品目の金額を合算する数式を設定
  var sumFormula = '=';
  for (var s = 0; s < 10; s++) {
    var sumRow = 13 + (s * 2);
    sumFormula += (s > 0 ? '+' : '') + 'G' + sumRow;
  }
  sheet.getRange('G32').setFormula(sumFormula);
  sheet.getRange('G32').setNumberFormat('#,##0');
  sheet.getRange('G32').setFontWeight('bold');
  sheet.getRange('G32').setFontSize(12);
  sheet.getRange('G32').setBackground('#D9E2F3');
  sheet.getRange('E32:G32').setBorder(true, true, true, true, null, null);

  // --- 備考欄 ---
  sheet.getRange('B34').setValue('備考');
  sheet.getRange('B34').setFontWeight('bold');
  var remarksArea = sheet.getRange('C34:G36');
  remarksArea.merge();
  remarksArea.setBackground('#FFF2CC');
  remarksArea.setVerticalAlignment('top');
  remarksArea.setWrap(true);
  remarksArea.setBorder(true, true, true, true, null, null);

  // --- 罫線で入力欄を囲む ---
  sheet.getRange('B3:C3').setBorder(null, null, true, null, null, null);
  sheet.getRange('B5:C5').setBorder(null, null, true, null, null, null);
  sheet.getRange('B7:C7').setBorder(null, null, true, null, null, null);
  sheet.getRange('B9:C9').setBorder(null, null, true, null, null, null);

  // --- 色分けの凡例（初心者が入力欄と自動計算欄を区別できるように） ---
  var legendRow = 38;
  sheet.getRange('B' + legendRow + ':G' + legendRow).merge();
  sheet.getRange('B' + legendRow).setValue(
    '【色の説明】  ■ 黄色 = 入力してください　■ 緑色 = 自動入力（単価・単位・金額・支払期限）　■ 青色 = 合計表示'
  );
  sheet.getRange('B' + legendRow).setFontSize(9);
  sheet.getRange('B' + legendRow).setFontColor('#666666');
  sheet.getRange('B' + legendRow).setBackground('#F5F5F5');

  // 行を固定
  sheet.setFrozenRows(2);
}

/**
 * 履歴管理シートを設定する（内部関数）
 * @param {Sheet} sheet - 履歴管理シート
 */
function setupHistorySheet_(sheet) {
  // ヘッダー
  var headers = ['発行番号', '書類種別', '発行日', '顧客名', '合計金額', 'PDFリンク', '発行日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4472C4');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');

  // 列幅調整
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 160);

  // 行を固定
  sheet.setFrozenRows(1);
}
