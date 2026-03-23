// ============================================================
// pdf.gs - PDF出力・ファイル管理
// 請求書・見積書 自動生成ツール
// ============================================================

/**
 * 出力シートをPDFとしてGoogleドライブに保存する
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} docNumber - 発行番号
 * @param {string} docType - 書類種別（請求書 or 見積書）
 * @return {string} PDFファイルの共有リンク
 */
function exportToPDF(ss, docNumber, docType) {
  // スプレッドシートが渡されなかった場合はアクティブなものを使用
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  var outputSheet = ss.getSheetByName('出力シート');
  if (!outputSheet) {
    throw new Error('「出力シート」シートが見つかりません。');
  }

  // --- 出力シートのIDを取得 ---
  var sheetId = outputSheet.getSheetId();

  // --- PDF生成用のURLを組み立て ---
  // GASからスプレッドシートをPDFに変換するためのURLパラメータ
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'exportFormat=pdf' +       // PDF形式で出力
    '&format=pdf' +
    '&size=A4' +               // A4サイズ
    '&portrait=true' +         // 縦向き
    '&fitw=true' +             // 幅に合わせる
    '&gridlines=false' +       // グリッド線を非表示
    '&printtitle=false' +      // タイトルを非表示
    '&sheetnames=false' +      // シート名を非表示
    '&pagenum=false' +         // ページ番号を非表示
    '&fzr=false' +             // 固定行を非表示
    '&top_margin=0.50' +       // 上余白（インチ）
    '&bottom_margin=0.50' +    // 下余白
    '&left_margin=0.50' +      // 左余白
    '&right_margin=0.50' +     // 右余白
    '&gid=' + sheetId;         // 出力対象のシートID

  // --- 認証トークンを使ってPDFを取得 ---
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF生成に失敗しました。レスポンスコード: ' + response.getResponseCode());
  }

  var pdfBlob = response.getBlob();

  // --- ファイル名を設定 ---
  // 例: 請求書_INV-20250101-001.pdf
  var fileName = docType + '_' + docNumber + '.pdf';
  pdfBlob.setName(fileName);

  // --- 保存先フォルダを取得（なければ作成） ---
  var folder = getOrCreateFolder_('請求書フォルダ');

  // --- PDFをフォルダに保存 ---
  var file = folder.createFile(pdfBlob);

  // --- 共有設定（リンクを知っている人が閲覧可能） ---
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // --- PDFリンクとフォルダリンクの両方を返す ---
  return {
    fileUrl: file.getUrl(),
    folderUrl: folder.getUrl()
  };
}

/**
 * Googleドライブ内のフォルダを取得する。存在しなければ新規作成する（内部関数）
 * @param {string} folderName - フォルダ名
 * @return {Folder} フォルダオブジェクト
 */
function getOrCreateFolder_(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    // 既存フォルダが見つかった場合はそれを返す
    return folders.next();
  } else {
    // フォルダが存在しない場合は新規作成
    Logger.log('フォルダ「' + folderName + '」を新規作成しました。');
    return DriveApp.createFolder(folderName);
  }
}
