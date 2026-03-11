// ============================================================
// Code.gs - メイン処理
// [ツール名をここに記載]
// ============================================================
//
// 概要:
//   [このツールの概要を記載]
//
// 使い方:
//   1. [手順1]
//   2. [手順2]
//
// ============================================================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ツール名')
    .addItem('メイン機能', 'mainFunction')
    .addSeparator()
    .addItem('使い方を見る', 'showHelp')
    .addToUi();
}

/**
 * メイン処理
 * TODO: ツールのメイン機能を実装する
 */
function mainFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    // TODO: メイン処理を実装
    ui.alert('完了', '処理が完了しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラーが発生しました', e.message, ui.ButtonSet.OK);
    Logger.log('mainFunction エラー: ' + e.message + '\n' + e.stack);
  }
}

/**
 * 使い方を表示する
 */
function showHelp() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '使い方',
    '【ツール名 使い方】\n\n' +
    '1. [手順1]\n' +
    '2. [手順2]\n' +
    '3. [手順3]',
    ui.ButtonSet.OK
  );
}
