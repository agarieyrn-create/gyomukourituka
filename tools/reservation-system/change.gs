// ============================================================
// change.gs - 予約変更処理
// GAS 予約管理ツール
// ============================================================

/**
 * 予約変更処理のメイン関数
 * 重複チェック・予約更新・カレンダー更新・履歴記録・メール送信を行う
 * @param {string} reservationId - 予約ID
 * @param {string} newDate - 変更後の日付（yyyy-MM-dd）
 * @param {string} newTime - 変更後の時間（HH:mm）
 * @return {Object} 処理結果
 */
function processChange(reservationId, newDate, newTime) {
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
        break;
      }
    }

    if (!reservation) {
      throw new Error('予約が見つかりません。');
    }

    if (reservation.status === 'キャンセル') {
      throw new Error('キャンセル済みの予約は変更できません。');
    }

    // 変更先が休業日でないかチェック
    if (isHoliday(newDate)) {
      throw new Error('変更先の日付は休業日です。別の日付を選択してください。');
    }

    // メニューの所要時間を取得
    var menuList = getMenuList();
    var duration = 60;
    for (var m = 0; m < menuList.length; m++) {
      if (menuList[m].name === reservation.menuName) {
        duration = menuList[m].duration;
        break;
      }
    }

    // 重複チェック（自身を除外）
    if (checkDuplicate(newDate, newTime, duration, reservationId)) {
      throw new Error('変更先の時間帯はすでに予約が入っています。');
    }

    // 元の予約日時を保存
    var originalDateTime = reservation.date + ' ' + reservation.time;

    // 1. 予約一覧を更新
    updateReservation(reservationId, {
      date: newDate,
      time: newTime,
      status: '確定' // ステータスは確定のまま
    });

    // 2. カレンダーイベントを更新
    var updatedReservation = {
      reservationId: reservationId,
      date: newDate,
      time: newTime,
      name: reservation.name,
      phone: reservation.phone,
      email: reservation.email,
      menuName: reservation.menuName,
      price: reservation.price,
      duration: duration,
      remarks: ''
    };

    if (reservation.eventId) {
      updateCalendarEvent(reservation.eventId, updatedReservation);
    }

    // 3. 履歴に記録
    saveToHistory({
      reservationId: reservationId,
      name: reservation.name,
      originalDateTime: originalDateTime,
      type: '変更',
      newDateTime: newDate + ' ' + newTime
    });

    // 4. 変更メール送信
    updatedReservation.reservationId = reservationId;
    sendChangeMail(updatedReservation);

    // 5. 管理者通知
    sendAdminNotification(updatedReservation, '変更');

    // 6. ダッシュボード更新
    updateDashboard();

    return {
      reservationId: reservationId,
      name: reservation.name,
      oldDate: reservation.date,
      oldTime: reservation.time,
      newDate: newDate,
      newTime: newTime,
      status: '確定'
    };

  } finally {
    lock.releaseLock();
  }
}
