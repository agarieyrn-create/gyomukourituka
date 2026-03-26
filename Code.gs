// ============================================================
// Code.gs - メイン処理
// 請求書・見積書 自動生成ツール
// ============================================================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 * この関数はGASの「onOpen」トリガーにより自動実行されます
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 書類生成ツール')
    .addItem('書類を生成する', 'generateDocument')
    .addItem('フォームをリセット', 'clearForm')
    .addSeparator()
    .addItem('初回セットアップ', 'setupProperties')
    .addItem('使い方を見る', 'showHelp')
    .addToUi();
}

// ============================================================
// メイン関数：書類生成
// ============================================================

/**
 * 入力フォームの内容を出力シートに転記し、PDF生成・履歴記録を実行する
 * 【書類を生成する】ボタンから呼び出されます
 */
function generateDocument() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    // --- キャッシュを初期化（前回実行の残留データをクリア） ---
    itemMasterCache_ = null;

    // --- 入力フォームからデータを取得 ---
    var formData = getFormData_(ss);

    // --- 入力値のバリデーション ---
    var validationError = validateFormData_(formData);
    if (validationError) {
      ui.alert('入力エラー', validationError, ui.ButtonSet.OK);
      return;
    }

    // --- 顧客マスタから顧客情報を取得 ---
    var customerData = getCustomerData(formData.customerName);
    if (!customerData) {
      ui.alert('エラー', '顧客情報が見つかりません: ' + formData.customerName, ui.ButtonSet.OK);
      return;
    }

    // --- 品目データを整理 ---
    var items = getItemsFromForm_(formData);
    if (items.length === 0) {
      ui.alert('入力エラー', '品目が1つも入力されていません。', ui.ButtonSet.OK);
      return;
    }

    // --- 発行番号を自動採番 ---
    var docNumber = getNextDocNumber(formData.docType);

    // --- 自社情報をスクリプトプロパティから取得 ---
    var companyInfo = getCompanyInfo_();
    if (!companyInfo.companyName) {
      ui.alert('セットアップ未完了', '先に「初回セットアップ」を実行してください。', ui.ButtonSet.OK);
      return;
    }

    // --- 出力シートにデータを書き込み ---
    writeToOutputSheet_(ss, formData, customerData, items, docNumber, companyInfo);

    // --- PDF出力 ---
    var pdfResult = exportToPDF(ss, docNumber, formData.docType);

    // --- 履歴管理に記録 ---
    var totalAmount = calculateTotal_(items);
    saveToHistory(docNumber, formData, totalAmount, pdfResult.fileUrl);

    // --- 完了メッセージ（PDFリンク＆フォルダリンク付き） ---
    ui.alert(
      '生成完了',
      formData.docType + 'を生成しました！\n\n' +
      '発行番号: ' + docNumber + '\n' +
      '合計金額: ¥' + formatNumber_(totalAmount) + '\n\n' +
      '【PDFファイル】\n' + pdfResult.fileUrl + '\n\n' +
      '【保存フォルダ（Googleドライブ）】\n' + pdfResult.folderUrl + '\n\n' +
      '※ 上記リンクをコピーしてブラウザで開くと確認できます。\n' +
      '※ 「履歴管理」シートからもワンクリックでPDFを開けます。',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('エラーが発生しました', e.message, ui.ButtonSet.OK);
    Logger.log('generateDocument エラー: ' + e.message + '\n' + e.stack);
  }
}

// ============================================================
// 入力フォーム関連の内部関数
// ============================================================

/**
 * 入力フォームシートからデータを取得する（内部関数）
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @return {Object} フォームデータ
 */
function getFormData_(ss) {
  var sheet = ss.getSheetByName('入力フォーム');
  if (!sheet) throw new Error('「入力フォーム」シートが見つかりません。');

  var data = {
    docType: sheet.getRange('C3').getValue(),         // 書類種別
    customerName: sheet.getRange('C5').getValue(),     // 顧客名
    issueDate: sheet.getRange('C7').getValue(),        // 発行日
    paymentDeadline: sheet.getRange('C9').getValue(),  // 支払期限
    remarks: sheet.getRange('C34').getValue(),         // 備考欄
    items: []
  };

  // 品目1〜10を取得（B13行目から開始、各品目は2行使用：品名行と数量行）
  for (var i = 0; i < 10; i++) {
    var row = 13 + (i * 2);
    var itemName = sheet.getRange('C' + row).getValue();       // 品目名（ドロップダウン）
    var quantity = sheet.getRange('E' + row).getValue();       // 数量
    var amount = sheet.getRange('G' + row).getValue();         // 金額（自動計算）

    data.items.push({
      name: itemName,
      quantity: quantity,
      amount: amount
    });
  }

  return data;
}

/**
 * フォームデータのバリデーション（内部関数）
 * @param {Object} formData - フォームデータ
 * @return {string|null} エラーメッセージ（問題なければnull）
 */
function validateFormData_(formData) {
  if (!formData.docType) return '書類種別を選択してください。';
  if (!formData.customerName) return '顧客名を選択してください。';
  if (!formData.issueDate) return '発行日を入力してください。';
  if (formData.docType === '請求書' && !formData.paymentDeadline) {
    return '請求書の場合、支払期限を入力してください。';
  }
  return null;
}

/**
 * フォームデータから有効な品目リストを取得する（内部関数）
 * @param {Object} formData - フォームデータ
 * @return {Array} 品目リスト
 */
function getItemsFromForm_(formData) {
  var items = [];
  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    if (item.name && item.quantity > 0) {
      // 品目マスタから単価と税区分を取得
      var masterItem = getItemMasterData_(item.name);
      if (masterItem) {
        items.push({
          name: item.name,
          unitPrice: masterItem.unitPrice,
          unit: masterItem.unit,
          taxCategory: masterItem.taxCategory,
          quantity: item.quantity,
          amount: masterItem.unitPrice * item.quantity
        });
      }
    }
  }
  return items;
}

// 品目マスタのキャッシュ（1回の書類生成中に同シートを何度も読まないようにする）
var itemMasterCache_ = null;

/**
 * 品目マスタから品目情報を取得する（内部関数）
 * @param {string} itemName - 品名
 * @return {Object|null} 品目情報
 */
function getItemMasterData_(itemName) {
  if (!itemMasterCache_) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('品目マスタ');
    if (!sheet) return null;
    itemMasterCache_ = sheet.getDataRange().getValues();
  }

  var data = itemMasterCache_;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === itemName) {
      var taxCategory = data[i][4];
      if (taxCategory !== '10%' && taxCategory !== '8%' && taxCategory !== '非課税') {
        throw new Error(
          '品目「' + itemName + '」の税区分が不正です: "' + taxCategory + '"\n' +
          '品目マスタの税区分は「10%」「8%」「非課税」のいずれかを指定してください。'
        );
      }
      return {
        itemId: data[i][0],
        name: data[i][1],
        unitPrice: data[i][2],
        unit: data[i][3],
        taxCategory: taxCategory
      };
    }
  }
  return null;
}

// ============================================================
// 出力シートへの書き込み
// ============================================================

/**
 * 出力シートにデータを書き込む（内部関数）
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Object} formData - フォームデータ
 * @param {Object} customerData - 顧客データ
 * @param {Array} items - 品目リスト
 * @param {string} docNumber - 発行番号
 * @param {Object} companyInfo - 自社情報
 */
function writeToOutputSheet_(ss, formData, customerData, items, docNumber, companyInfo) {
  var sheet = ss.getSheetByName('出力シート');
  if (!sheet) throw new Error('「出力シート」シートが見つかりません。');

  // --- シートを初期化 ---
  sheet.clear();
  sheet.clearFormats();

  // --- A4縦の印刷設定 ---
  sheet.setColumnWidth(1, 30);   // A列（余白）
  sheet.setColumnWidth(2, 60);   // B列
  sheet.setColumnWidth(3, 180);  // C列（品名）
  sheet.setColumnWidth(4, 60);   // D列（数量）
  sheet.setColumnWidth(5, 60);   // E列（単位）
  sheet.setColumnWidth(6, 100);  // F列（単価）
  sheet.setColumnWidth(7, 110);  // G列（金額）
  sheet.setColumnWidth(8, 30);   // H列（余白）

  // --- タイトル ---
  var titleCell = sheet.getRange('B2:G2');
  titleCell.merge();
  titleCell.setValue(formData.docType);
  titleCell.setFontSize(22);
  titleCell.setFontWeight('bold');
  titleCell.setHorizontalAlignment('center');

  // --- 発行番号 ---
  sheet.getRange('B4').setValue('発行番号:');
  sheet.getRange('C4').setValue(docNumber);
  sheet.getRange('C4').setFontWeight('bold');

  // --- 発行日 ---
  sheet.getRange('F4').setValue('発行日:');
  var issueDateFormatted = Utilities.formatDate(
    new Date(formData.issueDate), 'Asia/Tokyo', 'yyyy年MM月dd日'
  );
  sheet.getRange('G4').setValue(issueDateFormatted);
  sheet.getRange('G4').setHorizontalAlignment('right');

  // --- 宛先（顧客情報） ---
  sheet.getRange('B7').setValue(customerData.companyName + ' 御中');
  sheet.getRange('B7').setFontSize(14);
  sheet.getRange('B7').setFontWeight('bold');
  sheet.getRange('B8').setValue(customerData.contactPerson + ' 様');
  sheet.getRange('B9').setValue('〒' + customerData.zipCode);
  sheet.getRange('B10').setValue(customerData.address);

  // --- 自社情報（右側） ---
  sheet.getRange('F7').setValue(companyInfo.companyName);
  sheet.getRange('F7').setFontWeight('bold');
  sheet.getRange('F8').setValue('〒' + companyInfo.zipCode);
  sheet.getRange('F9').setValue(companyInfo.address);
  sheet.getRange('F10').setValue('TEL: ' + companyInfo.phone);
  sheet.getRange('F11').setValue('Email: ' + companyInfo.email);

  // --- 支払期限（請求書のみ） ---
  if (formData.docType === '請求書' && formData.paymentDeadline) {
    sheet.getRange('B12').setValue('お支払期限:');
    sheet.getRange('B12').setFontWeight('bold');
    var deadlineFormatted = Utilities.formatDate(
      new Date(formData.paymentDeadline), 'Asia/Tokyo', 'yyyy年MM月dd日'
    );
    sheet.getRange('C12').setValue(deadlineFormatted);
  }

  // --- 合計金額を先に表示 ---
  var totals = calculateTaxBreakdown_(items);
  sheet.getRange('B14').setValue('ご' + (formData.docType === '請求書' ? '請求' : '見積') + '金額');
  sheet.getRange('B14').setFontWeight('bold');
  var totalCell = sheet.getRange('C14:D14');
  totalCell.merge();
  totalCell.setValue('¥' + formatNumber_(totals.grandTotal) + '-（税込）');
  totalCell.setFontSize(16);
  totalCell.setFontWeight('bold');

  // --- 明細テーブルヘッダー ---
  var headerRow = 17;
  var headers = ['', 'No.', '品名', '数量', '単位', '単価', '金額'];
  for (var h = 0; h < headers.length; h++) {
    var cell = sheet.getRange(headerRow, h + 1);
    cell.setValue(headers[h]);
    if (h >= 1) {
      cell.setBackground('#4472C4');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      cell.setHorizontalAlignment('center');
    }
  }

  // --- 明細データ ---
  for (var i = 0; i < items.length; i++) {
    var row = headerRow + 1 + i;
    sheet.getRange(row, 2).setValue(i + 1);
    sheet.getRange(row, 2).setHorizontalAlignment('center');
    sheet.getRange(row, 3).setValue(items[i].name);
    sheet.getRange(row, 4).setValue(items[i].quantity);
    sheet.getRange(row, 4).setHorizontalAlignment('right');
    sheet.getRange(row, 5).setValue(items[i].unit);
    sheet.getRange(row, 5).setHorizontalAlignment('center');
    sheet.getRange(row, 6).setValue('¥' + formatNumber_(items[i].unitPrice));
    sheet.getRange(row, 6).setHorizontalAlignment('right');
    sheet.getRange(row, 7).setValue('¥' + formatNumber_(items[i].amount));
    sheet.getRange(row, 7).setHorizontalAlignment('right');

    // 偶数行に背景色
    if (i % 2 === 1) {
      sheet.getRange(row, 2, 1, 6).setBackground('#D6E4F0');
    }
  }

  // --- 明細テーブルの罫線 ---
  var tableRange = sheet.getRange(headerRow, 2, items.length + 1, 6);
  tableRange.setBorder(true, true, true, true, true, true);

  // --- 小計・税額・合計 ---
  var summaryStartRow = headerRow + items.length + 2;

  sheet.getRange(summaryStartRow, 5, 1, 2).merge();
  sheet.getRange(summaryStartRow, 5).setValue('小計');
  sheet.getRange(summaryStartRow, 5).setHorizontalAlignment('right');
  sheet.getRange(summaryStartRow, 5).setFontWeight('bold');
  sheet.getRange(summaryStartRow, 7).setValue('¥' + formatNumber_(totals.subtotal));
  sheet.getRange(summaryStartRow, 7).setHorizontalAlignment('right');

  var taxRow = summaryStartRow + 1;
  if (totals.tax10 > 0) {
    sheet.getRange(taxRow, 5, 1, 2).merge();
    sheet.getRange(taxRow, 5).setValue('消費税（10%）');
    sheet.getRange(taxRow, 5).setHorizontalAlignment('right');
    sheet.getRange(taxRow, 7).setValue('¥' + formatNumber_(totals.tax10));
    sheet.getRange(taxRow, 7).setHorizontalAlignment('right');
    taxRow++;
  }
  if (totals.tax8 > 0) {
    sheet.getRange(taxRow, 5, 1, 2).merge();
    sheet.getRange(taxRow, 5).setValue('消費税（8%）');
    sheet.getRange(taxRow, 5).setHorizontalAlignment('right');
    sheet.getRange(taxRow, 7).setValue('¥' + formatNumber_(totals.tax8));
    sheet.getRange(taxRow, 7).setHorizontalAlignment('right');
    taxRow++;
  }

  // 合計行
  sheet.getRange(taxRow, 5, 1, 2).merge();
  sheet.getRange(taxRow, 5).setValue('合計（税込）');
  sheet.getRange(taxRow, 5).setHorizontalAlignment('right');
  sheet.getRange(taxRow, 5).setFontWeight('bold');
  sheet.getRange(taxRow, 7).setValue('¥' + formatNumber_(totals.grandTotal));
  sheet.getRange(taxRow, 7).setHorizontalAlignment('right');
  sheet.getRange(taxRow, 7).setFontWeight('bold');
  sheet.getRange(taxRow, 7).setFontSize(13);

  // 合計の下に罫線
  sheet.getRange(summaryStartRow, 5, taxRow - summaryStartRow + 1, 3)
    .setBorder(true, true, true, true, true, true);

  // --- 備考欄 ---
  var remarksRow = taxRow + 3;
  sheet.getRange(remarksRow, 2).setValue('備考');
  sheet.getRange(remarksRow, 2).setFontWeight('bold');
  var remarksCell = sheet.getRange(remarksRow + 1, 2, 3, 6);
  remarksCell.merge();
  remarksCell.setValue(formData.remarks || '');
  remarksCell.setVerticalAlignment('top');
  remarksCell.setWrap(true);
  remarksCell.setBorder(true, true, true, true, null, null);

  // --- 振込先情報（請求書のみ） ---
  if (formData.docType === '請求書' && companyInfo.bankInfo) {
    var bankRow = remarksRow + 5;
    sheet.getRange(bankRow, 2).setValue('お振込先');
    sheet.getRange(bankRow, 2).setFontWeight('bold');
    var bankCell = sheet.getRange(bankRow + 1, 2, 2, 6);
    bankCell.merge();
    bankCell.setValue(companyInfo.bankInfo);
    bankCell.setVerticalAlignment('top');
    bankCell.setWrap(true);
    bankCell.setBorder(true, true, true, true, null, null);
    bankCell.setBackground('#FFF2CC');
  }

  // --- 出力シートをアクティブにする ---
  ss.setActiveSheet(sheet);
}

// ============================================================
// 計算関連の内部関数
// ============================================================

/**
 * 税区分ごとの内訳を計算する（内部関数）
 * @param {Array} items - 品目リスト
 * @return {Object} 小計・税額・合計
 */
function calculateTaxBreakdown_(items) {
  var subtotal10 = 0; // 10%対象の小計
  var subtotal8 = 0;  // 8%対象の小計
  var subtotalNon = 0; // 非課税の小計

  for (var i = 0; i < items.length; i++) {
    var amount = items[i].amount;
    var taxCat = items[i].taxCategory;

    if (taxCat === '10%') {
      subtotal10 += amount;
    } else if (taxCat === '8%') {
      subtotal8 += amount;
    } else { // '非課税'
      subtotalNon += amount;
    }
  }

  var tax10 = Math.floor(subtotal10 * 0.10);
  var tax8 = Math.floor(subtotal8 * 0.08);
  var subtotal = subtotal10 + subtotal8 + subtotalNon;
  var grandTotal = subtotal + tax10 + tax8;

  return {
    subtotal: subtotal,
    tax10: tax10,
    tax8: tax8,
    grandTotal: grandTotal
  };
}

/**
 * 品目リストから合計金額（税込）を計算する
 * @param {Array} items - 品目リスト
 * @return {number} 合計金額
 */
function calculateTotal_(items) {
  var totals = calculateTaxBreakdown_(items);
  return totals.grandTotal;
}

// ============================================================
// ユーティリティ関数
// ============================================================

/**
 * 数値を3桁カンマ区切りにフォーマットする（内部関数）
 * @param {number} num - 数値
 * @return {string} カンマ区切り文字列
 */
function formatNumber_(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * 自社情報をスクリプトプロパティから取得する（内部関数）
 * @return {Object} 自社情報
 */
function getCompanyInfo_() {
  var props = PropertiesService.getScriptProperties();
  return {
    companyName: props.getProperty('COMPANY_NAME') || '',
    zipCode: props.getProperty('ZIP_CODE') || '',
    address: props.getProperty('ADDRESS') || '',
    phone: props.getProperty('PHONE') || '',
    email: props.getProperty('EMAIL') || '',
    bankInfo: props.getProperty('BANK_INFO') || ''
  };
}

/**
 * 入力フォームをリセットする
 * カスタムメニューまたはボタンから呼び出されます
 */
function clearForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('入力フォーム');
  if (!sheet) return;

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'フォームリセット',
    '入力フォームの内容をすべてクリアしますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // 書類種別クリア
  sheet.getRange('C3').clearContent();
  // 顧客名クリア
  sheet.getRange('C5').clearContent();
  // 発行日を今日の日付にリセット
  sheet.getRange('C7').setValue(new Date());
  // 支払期限を自動入力式にリセット（発行日+30日）
  sheet.getRange('C9').setFormula('=IF(C7="","",C7+30)');

  // 品目1〜10をクリア
  for (var i = 0; i < 10; i++) {
    var row = 13 + (i * 2);
    sheet.getRange('C' + row).clearContent();     // 品名
    sheet.getRange('E' + row).clearContent();     // 数量
    // 金額は数式なのでクリアしない
  }

  // 備考欄クリア
  sheet.getRange('C34').clearContent();

  ui.alert('完了', 'フォームをリセットしました。', ui.ButtonSet.OK);
}

/**
 * 使い方を表示する
 */
function showHelp() {
  var ui = SpreadsheetApp.getUi();
  var helpText =
    '【請求書・見積書 自動生成ツール 使い方】\n\n' +
    '■ 初回セットアップ\n' +
    '  1. メニュー「📄 書類生成ツール」→「初回セットアップ」\n' +
    '  2. 自社の会社名・住所・振込先などを入力\n\n' +
    '■ 書類の作成手順\n' +
    '  1. 「入力フォーム」シートを開く\n' +
    '  2. 書類種別（請求書 or 見積書）を選択\n' +
    '  3. 顧客名をドロップダウンから選択\n' +
    '  4. 発行日を入力（請求書の場合は支払期限も）\n' +
    '  5. 品目をドロップダウンから選択し、数量を入力\n' +
    '  6. 必要に応じて備考欄に記入\n' +
    '  7. メニュー「📄 書類生成ツール」→「書類を生成する」\n\n' +
    '■ 生成後\n' +
    '  ・「出力シート」に書類が表示されます\n' +
    '  ・PDFがGoogleドライブの「請求書フォルダ」に保存されます\n' +
    '  ・「履歴管理」シートに記録が追加されます\n\n' +
    '■ マスタデータの管理\n' +
    '  ・「顧客マスタ」「品目マスタ」シートで顧客・品目を追加・編集できます\n' +
    '  ・追加した顧客・品目は入力フォームのドロップダウンに自動反映されます';

  ui.alert('使い方', helpText, ui.ButtonSet.OK);
}
