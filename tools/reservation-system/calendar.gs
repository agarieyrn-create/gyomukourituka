// ============================================================
// calendar.gs - カレンダー関連処理
// GAS 予約管理ツール
// ============================================================

/**
 * 月全体の空き状況を返す
 * 定休日・臨時休業日の両方を考慮する
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 * @return {Array} 日ごとの空き状況配列
 */
function getMonthAvailability(year, month) {
  var settings = getSettings();
  var interval = parseInt(settings['予約間隔（分）'] || '60', 10);
  var maxSlots = parseInt(settings['同時受付数'] || '1', 10);
  var maxBookingDays = parseInt(settings['予約受付期間（何日先まで）'] || '30', 10);

  // 月の日数を取得
  var daysInMonth = new Date(year, month, 0).getDate();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // 予約受付の最終日
  var maxDate = new Date(today);
  maxDate.setDate(maxDate.getDate() + maxBookingDays);

  // 休業日リストを取得
  var closedDates = getClosedDates(year, month);

  // 全予約を取得（当月分）
  var reservations = getMonthReservations_(year, month);

  // 1日あたりの最大枠数を計算
  var openTime = settings['営業開始時間'] || '10:00';
  var closeTime = settings['営業終了時間'] || '19:00';
  var openMinutes = timeToMinutes_(openTime);
  var closeMinutes = timeToMinutes_(closeTime);
  var totalSlotsPerDay = Math.floor((closeMinutes - openMinutes) / interval) * maxSlots;

  var result = [];

  for (var day = 1; day <= daysInMonth; day++) {
    var date = new Date(year, month - 1, day);
    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');

    var dayInfo = {
      date: dateStr,
      day: day,
      dayOfWeek: date.getDay(),
      status: 'available' // available, full, closed, past
    };

    // 過去日チェック
    if (date < today) {
      dayInfo.status = 'past';
      result.push(dayInfo);
      continue;
    }

    // 予約受付期間外チェック
    if (date > maxDate) {
      dayInfo.status = 'past';
      result.push(dayInfo);
      continue;
    }

    // 休業日チェック
    if (closedDates.indexOf(dateStr) !== -1) {
      dayInfo.status = 'closed';
      result.push(dayInfo);
      continue;
    }

    // 定休日チェック
    if (isRegularHoliday_(date, settings)) {
      dayInfo.status = 'closed';
      result.push(dayInfo);
      continue;
    }

    // 予約済み枠数を計算
    var bookedCount = 0;
    for (var r = 0; r < reservations.length; r++) {
      if (reservations[r].date === dateStr && reservations[r].status === '確定') {
        bookedCount++;
      }
    }

    // 満枠チェック
    if (bookedCount >= totalSlotsPerDay) {
      dayInfo.status = 'full';
    }

    result.push(dayInfo);
  }

  return result;
}

/**
 * 指定日の時間枠ごとの空き状況を返す
 * @param {string} dateStr - 日付文字列（yyyy-MM-dd）
 * @return {Array} 時間枠の空き状況配列
 */
function getDaySlots(dateStr) {
  var settings = getSettings();
  var interval = parseInt(settings['予約間隔（分）'] || '60', 10);
  var maxSlots = parseInt(settings['同時受付数'] || '1', 10);
  var openTime = settings['営業開始時間'] || '10:00';
  var closeTime = settings['営業終了時間'] || '19:00';

  var openMinutes = timeToMinutes_(openTime);
  var closeMinutes = timeToMinutes_(closeTime);

  // 指定日の確定済み予約を取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  var reservations = [];

  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowDate = Utilities.formatDate(new Date(data[i][2]), 'Asia/Tokyo', 'yyyy-MM-dd');
      if (rowDate === dateStr && data[i][9] === '確定') {
        reservations.push({
          time: data[i][3],
          menuName: data[i][7]
        });
      }
    }
  }

  // 現在時刻（当日の場合、過去の時間枠は選択不可にする）
  var now = new Date();
  var targetDate = new Date(dateStr + 'T00:00:00+09:00');
  var isToday = (now.toDateString() === targetDate.toDateString());

  var slots = [];
  for (var min = openMinutes; min < closeMinutes; min += interval) {
    var timeStr = minutesToTime_(min);
    var slotInfo = {
      time: timeStr,
      available: true
    };

    // 当日の過去時刻チェック
    if (isToday) {
      var nowMinutes = now.getHours() * 60 + now.getMinutes();
      if (min <= nowMinutes) {
        slotInfo.available = false;
        slots.push(slotInfo);
        continue;
      }
    }

    // この時間枠の予約数をカウント
    var count = 0;
    for (var r = 0; r < reservations.length; r++) {
      if (reservations[r].time === timeStr) {
        count++;
      }
    }

    if (count >= maxSlots) {
      slotInfo.available = false;
    }

    slots.push(slotInfo);
  }

  return slots;
}

/**
 * Googleカレンダーにイベントを登録しIDを返す
 * @param {Object} reservation - 予約データ
 * @return {string} カレンダーイベントID
 */
function registerCalendarEvent(reservation) {
  var settings = getSettings();
  var calendarId = settings['共通GoogleカレンダーID'];
  if (!calendarId) return '';

  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log('カレンダーが見つかりません: ' + calendarId);
      return '';
    }

    var startTime = new Date(reservation.date + 'T' + reservation.time + ':00+09:00');
    var endTime = new Date(startTime.getTime() + reservation.duration * 60 * 1000);

    var title = '【予約】' + reservation.name + ' / ' + reservation.menuName;
    var description =
      '予約ID: ' + reservation.reservationId + '\n' +
      '氏名: ' + reservation.name + '\n' +
      '電話: ' + reservation.phone + '\n' +
      'メール: ' + reservation.email + '\n' +
      'メニュー: ' + reservation.menuName + '\n' +
      '料金: ¥' + reservation.price.toLocaleString() + '\n' +
      '備考: ' + (reservation.remarks || 'なし');

    var event = calendar.createEvent(title, startTime, endTime, {
      description: description
    });

    return event.getId();

  } catch (e) {
    Logger.log('カレンダーイベント登録エラー: ' + e.message);
    return '';
  }
}

/**
 * カレンダーイベントを更新する
 * @param {string} eventId - カレンダーイベントID
 * @param {Object} reservation - 更新後の予約データ
 */
function updateCalendarEvent(eventId, reservation) {
  var settings = getSettings();
  var calendarId = settings['共通GoogleカレンダーID'];
  if (!calendarId || !eventId) return;

  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) return;

    var event = calendar.getEventById(eventId);
    if (!event) return;

    var startTime = new Date(reservation.date + 'T' + reservation.time + ':00+09:00');
    var endTime = new Date(startTime.getTime() + reservation.duration * 60 * 1000);

    event.setTime(startTime, endTime);

    var title = '【予約】' + reservation.name + ' / ' + reservation.menuName;
    event.setTitle(title);

    var description =
      '予約ID: ' + reservation.reservationId + '\n' +
      '氏名: ' + reservation.name + '\n' +
      '電話: ' + reservation.phone + '\n' +
      'メール: ' + reservation.email + '\n' +
      'メニュー: ' + reservation.menuName + '\n' +
      '料金: ¥' + reservation.price.toLocaleString() + '\n' +
      '備考: ' + (reservation.remarks || 'なし');

    event.setDescription(description);

  } catch (e) {
    Logger.log('カレンダーイベント更新エラー: ' + e.message);
  }
}

/**
 * カレンダーイベントを削除する
 * @param {string} eventId - カレンダーイベントID
 */
function deleteCalendarEvent(eventId) {
  var settings = getSettings();
  var calendarId = settings['共通GoogleカレンダーID'];
  if (!calendarId || !eventId) return;

  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) return;

    var event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
    }
  } catch (e) {
    Logger.log('カレンダーイベント削除エラー: ' + e.message);
  }
}

/**
 * 重複予約チェック
 * 同じ日時に同時受付数を超える予約がないか確認する
 * @param {string} dateStr - 日付文字列
 * @param {string} timeStr - 時間文字列
 * @param {number} duration - 所要時間（分）
 * @param {string|null} excludeId - 除外する予約ID（変更時）
 * @return {boolean} 重複している場合true
 */
function checkDuplicate(dateStr, timeStr, duration, excludeId) {
  var settings = getSettings();
  var maxSlots = parseInt(settings['同時受付数'] || '1', 10);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  if (!sheet || sheet.getLastRow() <= 1) return false;

  var data = sheet.getDataRange().getValues();
  var count = 0;

  // リクエスト予約の開始・終了時間（分）
  var reqStart = timeToMinutes_(timeStr);
  var reqEnd = reqStart + duration;

  for (var i = 1; i < data.length; i++) {
    // 除外ID（自身の予約）はスキップ
    if (excludeId && data[i][0] === excludeId) continue;
    // キャンセル済みはスキップ
    if (data[i][9] !== '確定') continue;

    var rowDate = Utilities.formatDate(new Date(data[i][2]), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (rowDate !== dateStr) continue;

    // 既存予約の開始・終了時間
    var existStart = timeToMinutes_(data[i][3]);
    // メニューの所要時間を取得（デフォルト60分）
    var existDuration = getMenuDuration_(data[i][7]) || 60;
    var existEnd = existStart + existDuration;

    // 時間が重なるかチェック
    if (reqStart < existEnd && reqEnd > existStart) {
      count++;
    }
  }

  return count >= maxSlots;
}

// ============================================================
// 内部ヘルパー関数
// ============================================================

/**
 * 時間文字列を分に変換する（内部関数）
 * @param {string} timeStr - "HH:mm"形式の時間文字列
 * @return {number} 分
 */
function timeToMinutes_(timeStr) {
  var parts = timeStr.split(':');
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

/**
 * 分を時間文字列に変換する（内部関数）
 * @param {number} minutes - 分
 * @return {string} "HH:mm"形式の時間文字列
 */
function minutesToTime_(minutes) {
  var h = Math.floor(minutes / 60);
  var m = minutes % 60;
  return (h < 10 ? '0' + h : h) + ':' + (m < 10 ? '0' + m : m);
}

/**
 * 指定月の全予約を取得する（内部関数）
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 * @return {Array} 予約データ配列
 */
function getMonthReservations_(year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var reservations = [];

  for (var i = 1; i < data.length; i++) {
    var rowDate = new Date(data[i][2]);
    if (rowDate.getFullYear() === year && (rowDate.getMonth() + 1) === month) {
      reservations.push({
        date: Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
        time: data[i][3],
        status: data[i][9]
      });
    }
  }

  return reservations;
}

/**
 * 定休日かどうかを判定する（内部関数）
 * @param {Date} date - 日付
 * @param {Object} settings - 設定値
 * @return {boolean} 定休日の場合true
 */
function isRegularHoliday_(date, settings) {
  var closedDaysStr = settings['定休日曜日'] || '';
  if (!closedDaysStr) return false;

  var dayNames = ['日曜', '月曜', '火曜', '水曜', '木曜', '金曜', '土曜'];
  var dayName = dayNames[date.getDay()];

  var closedDays = closedDaysStr.split(',').map(function(d) { return d.trim(); });
  return closedDays.indexOf(dayName) !== -1;
}

/**
 * メニュー名から所要時間を取得する（内部関数）
 * @param {string} menuName - メニュー名
 * @return {number} 所要時間（分）
 */
function getMenuDuration_(menuName) {
  var menuList = getMenuList();
  for (var i = 0; i < menuList.length; i++) {
    if (menuList[i].name === menuName) {
      return menuList[i].duration;
    }
  }
  return 60; // デフォルト60分
}
