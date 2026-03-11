// ============================================================
// reminder.gs - リマインダー送信処理
// GAS 予約管理ツール
// ============================================================

/**
 * リマインダー送信メイン関数
 * 毎日8時にトリガーで実行される
 * 設定シートの「リマインダー送信タイミング」に基づき対象予約にメール送信
 */
function sendReminders() {
  var settings = getSettings();
  var reminderTimingStr = settings['リマインダー送信タイミング'] || '1日前';

  // リマインダー送信タイミングを解析（例: "1日前" → [1], "1日前,2日前" → [1, 2]）
  var timings = reminderTimingStr.split(',').map(function(t) {
    var match = t.trim().match(/(\d+)/);
    return match ? parseInt(match[1], 10) : 1;
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  if (!sheet || sheet.getLastRow() <= 1) return;

  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  var sentCount = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = row[9];        // ステータス
    var reminderSent = row[12]; // リマインダー送信済みフラグ

    // 確定済み かつ 未送信のみ対象
    if (status !== '確定' || reminderSent === '送信済') continue;

    var reserveDate = new Date(row[2]);
    reserveDate.setHours(0, 0, 0, 0);

    // 予約日までの日数を計算
    var diffDays = Math.floor((reserveDate - today) / (1000 * 60 * 60 * 24));

    // 送信タイミングに該当するかチェック
    var shouldSend = false;
    for (var t = 0; t < timings.length; t++) {
      if (diffDays === timings[t]) {
        shouldSend = true;
        break;
      }
    }

    if (!shouldSend) continue;

    // リマインダーメール送信
    var reservation = {
      reservationId: row[0],
      date: Utilities.formatDate(reserveDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
      time: row[3],
      name: row[4],
      email: row[6],
      menuName: row[7]
    };

    sendReminderMail(reservation);

    // 送信済みフラグを更新
    sheet.getRange(i + 1, 13).setValue('送信済');
    sentCount++;
  }

  Logger.log('リマインダー送信完了: ' + sentCount + ' 件');
}

/**
 * 毎日8時のタイムトリガーを自動設定する
 * 既存のリマインダートリガーがあれば削除してから再設定する
 */
function setupReminderTrigger() {
  // 既存のリマインダートリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎日午前8時にトリガーを設定
  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('リマインダートリガーを設定しました（毎日8:00）');
}
