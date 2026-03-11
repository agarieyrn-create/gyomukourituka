// ============================================================
// cancel.gs - キャンセル処理
// GAS 予約管理ツール
// ============================================================

/**
 * キャンセル処理のメイン関数
 * 予約一覧のステータス更新・カレンダーイベント削除・履歴記録・メール送信を行う
 * @param {string} reservationId - 予約ID
 * @return {Object} 処理結果
 */
function processCancellation(reservationId) {
  // LockServiceで二重処理を防止
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('サーバーが混雑しています。しばらくしてから再度お試しください。');
  }

  try {
    // 予約データを取得
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('予約一覧');
    if (!sheet) throw new Error('予約一覧シートが見つかりません。');

    var data = sheet.getDataRange().getValues();
    var reservation = null;
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === reservationId) {
        reservation = {
          reservationId: data[i][0],
          date: Utilities.formatDate(new Date(data[i][2]), 'Asia/Tokyo', 'yyyy-MM-dd'),
          time: data[i][3],
          name: data[i][4],
          phone: data[i][5],
          email: data[i][6],
          menuName: data[i][7],
          price: data[i][8],
          status: data[i][9],
          eventId: data[i][11]
        };
        rowIndex = i;
        break;
      }
    }

    if (!reservation) {
      throw new Error('予約が見つかりません。');
    }

    if (reservation.status === 'キャンセル') {
      throw new Error('この予約はすでにキャンセル済みです。');
    }

    // 1. 予約一覧のステータスを更新
    cancelReservationStatus(reservationId);

    // 2. カレンダーイベントを削除
    if (reservation.eventId) {
      deleteCalendarEvent(reservation.eventId);
    }

    // 3. 履歴に記録
    saveToHistory({
      reservationId: reservationId,
      name: reservation.name,
      originalDateTime: reservation.date + ' ' + reservation.time,
      type: 'キャンセル',
      newDateTime: ''
    });

    // 4. キャンセルメール送信
    sendCancelMail(reservation);

    // 5. 管理者通知
    sendAdminNotification(reservation, 'キャンセル');

    // 6. ダッシュボード更新
    updateDashboard();

    return {
      reservationId: reservationId,
      name: reservation.name,
      date: reservation.date,
      time: reservation.time,
      status: 'キャンセル'
    };

  } finally {
    lock.releaseLock();
  }
}
