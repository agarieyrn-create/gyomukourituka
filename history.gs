// ============================================================
// history.gs - 履歴管理・採番
// 請求書・見積書 自動生成ツール
// ============================================================

/**
 * 発行情報を履歴管理シートに追記する
 * @param {string} docNumber - 発行番号
 * @param {Object} formData - フォームデータ
 * @param {number} totalAmount - 合計金額（税込）
 * @param {string} pdfUrl - PDFの共有リンク
 */
function saveToHistory(docNumber, formData, totalAmount, pdfUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('履歴管理');
  if (!sheet) throw new Error('「履歴管理」シートが見つかりません。');

  // 発行日をフォーマット
  var issueDateFormatted = Utilities.formatDate(
    new Date(formData.issueDate), 'Asia/Tokyo', 'yyyy/MM/dd'
  );

  // 現在日時を取得（発行日時として記録）
  var now = Utilities.formatDate(
    new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'
  );

  // 最終行の次の行にデータを追記
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;

  sheet.getRange(newRow, 1).setValue(docNumber);               // 発行番号
  sheet.getRange(newRow, 2).setValue(formData.docType);        // 書類種別
  sheet.getRange(newRow, 3).setValue(issueDateFormatted);      // 発行日
  sheet.getRange(newRow, 4).setValue(formData.customerName);   // 顧客名
  sheet.getRange(newRow, 5).setValue(totalAmount);             // 合計金額
  sheet.getRange(newRow, 5).setNumberFormat('#,##0');          // 金額フォーマット

  // PDFリンクをクリック可能なハイパーリンクとして設定
  // HYPERLINK関数を使い、セルをクリックするだけでPDFを開けるようにする
  var linkFormula = '=HYPERLINK("' + pdfUrl + '","PDFを開く")';
  sheet.getRange(newRow, 6).setFormula(linkFormula);           // PDFリンク
  sheet.getRange(newRow, 6).setFontColor('#1155CC');           // リンク色（青）

  sheet.getRange(newRow, 7).setValue(now);                     // 発行日時
}

/**
 * 書類種別に応じた次の発行番号を生成する
 * フォーマット: INV-YYYYMMDD-XXX（請求書）/ EST-YYYYMMDD-XXX（見積書）
 * @param {string} docType - 書類種別（"請求書" or "見積書"）
 * @return {string} 発行番号
 */
function getNextDocNumber(docType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('履歴管理');
  if (!sheet) throw new Error('「履歴管理」シートが見つかりません。');

  // プレフィックスを決定
  var prefix = (docType === '請求書') ? 'INV' : 'EST';

  // 今日の日付をYYYYMMDD形式に変換
  var today = new Date();
  var dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd');

  // 履歴データから同日・同種別の最大番号を検索
  var data = sheet.getDataRange().getValues();
  var maxNum = 0;
  var searchPrefix = prefix + '-' + dateStr + '-';

  for (var i = 1; i < data.length; i++) {
    var existingNumber = data[i][0].toString();
    if (existingNumber.indexOf(searchPrefix) === 0) {
      // 末尾3桁の番号を取得
      var numPart = parseInt(existingNumber.substring(searchPrefix.length), 10);
      if (numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }

  // 次の番号を生成（3桁ゼロ埋め）
  var nextNum = maxNum + 1;
  var numStr = ('000' + nextNum).slice(-3);

  return prefix + '-' + dateStr + '-' + numStr;
}

/**
 * 顧客マスタから顧客情報を取得する
 * @param {string} name - 会社名
 * @return {Object|null} 顧客情報オブジェクト
 */
function getCustomerData(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('顧客マスタ');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();

  // 1行目はヘッダーなのでスキップ
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      return {
        customerId: data[i][0],
        companyName: data[i][1],
        contactPerson: data[i][2],
        zipCode: data[i][3],
        address: data[i][4],
        email: data[i][5],
        phone: data[i][6]
      };
    }
  }

  return null;
}
