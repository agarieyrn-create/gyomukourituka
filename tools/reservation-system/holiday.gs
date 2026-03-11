// ============================================================
// holiday.gs - 休業日管理
// GAS 予約管理ツール
// ============================================================

/**
 * 指定日が定休日または臨時休業日かを返す
 * @param {Date|string} date - 日付
 * @return {boolean} 休業日の場合true
 */
function isHoliday(date) {
  if (typeof date === 'string') {
    date = new Date(date + 'T00:00:00+09:00');
  }

  var settings = getSettings();

  // 定休日チェック
  var closedDaysStr = settings['定休日曜日'] || '';
  if (closedDaysStr) {
    var dayNames = ['日曜', '月曜', '火曜', '水曜', '木曜', '金曜', '土曜'];
    var dayName = dayNames[date.getDay()];
    var closedDays = closedDaysStr.split(',').map(function(d) { return d.trim(); });
    if (closedDays.indexOf(dayName) !== -1) {
      return true;
    }
  }

  // 臨時休業日チェック
  var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
  var holidays = getHolidayList();
  return holidays.indexOf(dateStr) !== -1;
}

/**
 * 月内の全休業日リストを返す
 * 定休日と臨時休業日の両方を含む
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 * @return {Array} 休業日の日付文字列配列（yyyy-MM-dd）
 */
function getClosedDates(year, month) {
  var settings = getSettings();
  var daysInMonth = new Date(year, month, 0).getDate();
  var closedDates = [];

  // 定休日を追加
  var closedDaysStr = settings['定休日曜日'] || '';
  var closedDays = closedDaysStr ? closedDaysStr.split(',').map(function(d) { return d.trim(); }) : [];
  var dayNames = ['日曜', '月曜', '火曜', '水曜', '木曜', '金曜', '土曜'];

  for (var day = 1; day <= daysInMonth; day++) {
    var date = new Date(year, month - 1, day);
    var dayName = dayNames[date.getDay()];
    if (closedDays.indexOf(dayName) !== -1) {
      closedDates.push(Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd'));
    }
  }

  // 臨時休業日を追加
  var holidays = getHolidayList();
  for (var h = 0; h < holidays.length; h++) {
    var hDate = new Date(holidays[h] + 'T00:00:00+09:00');
    if (hDate.getFullYear() === year && (hDate.getMonth() + 1) === month) {
      if (closedDates.indexOf(holidays[h]) === -1) {
        closedDates.push(holidays[h]);
      }
    }
  }

  return closedDates;
}

/**
 * 臨時休業日をシートに追加するUI関数
 * スプレッドシートのカスタムメニューから呼び出す
 * @param {string} dateStr - 日付文字列（yyyy-MM-dd）省略時はダイアログで入力
 * @param {string} reason - 理由（任意）
 */
function addTemporaryHoliday(dateStr, reason) {
  var ui = SpreadsheetApp.getUi();

  // 日付が未指定の場合はダイアログで入力
  if (!dateStr) {
    var dateResponse = ui.prompt(
      '臨時休業日の追加',
      '休業日の日付を入力してください（例: 2026-04-01）',
      ui.ButtonSet.OK_CANCEL
    );
    if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
    dateStr = dateResponse.getResponseText().trim();
  }

  // 日付の妥当性チェック
  var dateRegex = /^\d{4}-\d{2}-\d{2}$/;
  if (!dateRegex.test(dateStr)) {
    ui.alert('エラー', '日付の形式が正しくありません（yyyy-MM-dd）', ui.ButtonSet.OK);
    return;
  }

  // 理由が未指定の場合はダイアログで入力
  if (reason === undefined || reason === null) {
    var reasonResponse = ui.prompt(
      '臨時休業日の追加',
      '理由を入力してください（任意・空白可）',
      ui.ButtonSet.OK_CANCEL
    );
    if (reasonResponse.getSelectedButton() !== ui.Button.OK) return;
    reason = reasonResponse.getResponseText().trim();
  }

  // スプレッドシートに追加
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('臨時休業日');
  if (!sheet) {
    ui.alert('エラー', '臨時休業日シートが見つかりません。', ui.ButtonSet.OK);
    return;
  }

  sheet.appendRow([new Date(dateStr), reason || '']);

  ui.alert('完了', dateStr + ' を臨時休業日に追加しました。', ui.ButtonSet.OK);
}
