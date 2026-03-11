// ============================================================
// sheet.gs - スプレッドシート操作
// GAS 予約管理ツール
// ============================================================

/**
 * 予約一覧に記録し予約IDを返す
 * @param {Object} data - 予約データ
 * @return {string} 予約ID（RSV-YYYYMMDD-XXX形式）
 */
function saveReservation(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  if (!sheet) throw new Error('予約一覧シートが見つかりません。');

  // 予約IDを自動採番
  var reservationId = generateReservationId_(data.date);

  // 受付日時
  var now = new Date();
  var receiveDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  // 新しい行にデータを追加
  sheet.appendRow([
    reservationId,          // A: 予約ID
    receiveDate,            // B: 受付日時
    new Date(data.date),    // C: 予約日
    data.time,              // D: 予約時間
    data.name,              // E: 氏名
    data.phone,             // F: 電話番号
    data.email,             // G: メールアドレス
    data.menuName,          // H: メニュー
    data.price,             // I: 料金
    '確定',                  // J: ステータス
    data.remarks || '',     // K: 備考
    '',                     // L: カレンダーイベントID
    ''                      // M: リマインダー送信済みフラグ
  ]);

  return reservationId;
}

/**
 * 予約内容を更新する
 * @param {string} reservationId - 予約ID
 * @param {Object} newData - 更新データ
 */
function updateReservation(reservationId, newData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  if (!sheet) throw new Error('予約一覧シートが見つかりません。');

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      var row = i + 1;
      // 予約日を更新
      if (newData.date) {
        sheet.getRange(row, 3).setValue(new Date(newData.date));
      }
      // 予約時間を更新
      if (newData.time) {
        sheet.getRange(row, 4).setValue(newData.time);
      }
      // ステータスを更新
      if (newData.status) {
        sheet.getRange(row, 10).setValue(newData.status);
      }
      // リマインダーフラグをリセット（日時変更の場合）
      if (newData.date || newData.time) {
        sheet.getRange(row, 13).setValue('');
      }
      return;
    }
  }

  throw new Error('予約が見つかりません: ' + reservationId);
}

/**
 * ステータスをキャンセルに更新する
 * @param {string} reservationId - 予約ID
 */
function cancelReservationStatus(reservationId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  if (!sheet) throw new Error('予約一覧シートが見つかりません。');

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      sheet.getRange(i + 1, 10).setValue('キャンセル');
      return;
    }
  }

  throw new Error('予約が見つかりません: ' + reservationId);
}

/**
 * 設定シートから全設定値をオブジェクトで返す
 * @return {Object} キー・バリュー形式の設定値
 */
function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('設定');
  if (!sheet) throw new Error('設定シートが見つかりません。');

  var data = sheet.getDataRange().getValues();
  var settings = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      settings[data[i][0]] = data[i][1];
    }
  }

  return settings;
}

/**
 * メニューマスタを配列で返す
 * @return {Array} メニュー配列
 */
function getMenuList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('メニューマスタ');
  if (!sheet) throw new Error('メニューマスタシートが見つかりません。');

  var data = sheet.getDataRange().getValues();
  var menuList = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      menuList.push({
        menuId: data[i][0],
        name: data[i][1],
        duration: data[i][2],
        price: data[i][3]
      });
    }
  }

  return menuList;
}

/**
 * 臨時休業日シートから休業日一覧を返す
 * @return {Array} 休業日の日付文字列配列
 */
function getHolidayList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('臨時休業日');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var holidays = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var dateStr = Utilities.formatDate(new Date(data[i][0]), 'Asia/Tokyo', 'yyyy-MM-dd');
      holidays.push(dateStr);
    }
  }

  return holidays;
}

/**
 * キャンセル・変更履歴に記録する
 * @param {Object} data - 履歴データ
 */
function saveToHistory(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('キャンセル・変更履歴');
  if (!sheet) {
    Logger.log('キャンセル・変更履歴シートが見つかりません。');
    return;
  }

  var now = new Date();
  var processDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  sheet.appendRow([
    processDate,            // A: 処理日時
    data.reservationId,     // B: 予約ID
    data.name,              // C: 氏名
    data.originalDateTime,  // D: 元の予約日時
    data.type,              // E: 処理種別
    data.newDateTime || ''  // F: 変更後日時
  ]);
}

/**
 * ダッシュボードシートを最新データで更新する
 */
function updateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('ダッシュボード');
  if (!dashboard) return;

  var reservationSheet = ss.getSheetByName('予約一覧');
  if (!reservationSheet || reservationSheet.getLastRow() <= 1) {
    dashboard.getRange('A1').setValue('ダッシュボード（最終更新: ' +
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') + '）');
    return;
  }

  var data = reservationSheet.getDataRange().getValues();
  var now = new Date();
  var today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

  // --- 今日の予約一覧 ---
  var todayReservations = [];
  // --- 今週の集計 ---
  var weekStart = new Date(now);
  weekStart.setDate(weekStart.getDate() - weekStart.getDay());
  weekStart.setHours(0, 0, 0, 0);
  var weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);
  weekEnd.setHours(23, 59, 59, 999);

  var weekCount = 0;
  var weekSales = 0;

  // --- 今月の集計 ---
  var currentMonth = now.getMonth();
  var currentYear = now.getFullYear();
  var monthCount = 0;
  var monthSales = 0;

  // --- ステータス別件数 ---
  var statusCount = { '確定': 0, 'キャンセル': 0, '変更': 0 };

  // --- 直近7日間の予約件数 ---
  var last7days = {};
  for (var d = 6; d >= 0; d--) {
    var pastDate = new Date(now);
    pastDate.setDate(pastDate.getDate() - d);
    var key = Utilities.formatDate(pastDate, 'Asia/Tokyo', 'MM/dd');
    last7days[key] = 0;
  }

  // --- 次の予約 ---
  var nextReservation = null;
  var nextReservationTime = null;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var reserveDate = new Date(row[2]);
    var reserveDateStr = Utilities.formatDate(reserveDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    var status = row[9];

    // ステータス集計
    if (statusCount.hasOwnProperty(status)) {
      statusCount[status]++;
    }

    // 確定予約のみ集計
    if (status !== '確定') continue;

    // 今日の予約
    if (reserveDateStr === today) {
      todayReservations.push({
        time: row[3],
        name: row[4],
        menu: row[7]
      });
    }

    // 今週の集計
    if (reserveDate >= weekStart && reserveDate <= weekEnd) {
      weekCount++;
      weekSales += (row[8] || 0);
    }

    // 今月の集計
    if (reserveDate.getFullYear() === currentYear && reserveDate.getMonth() === currentMonth) {
      monthCount++;
      monthSales += (row[8] || 0);
    }

    // 直近7日間
    var dayKey = Utilities.formatDate(reserveDate, 'Asia/Tokyo', 'MM/dd');
    if (last7days.hasOwnProperty(dayKey)) {
      last7days[dayKey]++;
    }

    // 次の予約を探す
    var reserveDateTime = new Date(reserveDateStr + 'T' + row[3] + ':00+09:00');
    if (reserveDateTime > now) {
      if (!nextReservationTime || reserveDateTime < nextReservationTime) {
        nextReservationTime = reserveDateTime;
        nextReservation = {
          date: reserveDateStr,
          time: row[3],
          name: row[4],
          menu: row[7]
        };
      }
    }
  }

  // 今日の予約を時間順にソート
  todayReservations.sort(function(a, b) {
    return a.time.localeCompare(b.time);
  });

  // --- ダッシュボードに書き込み ---
  dashboard.clear();

  var r = 1;
  dashboard.getRange(r, 1).setValue('ダッシュボード');
  dashboard.getRange(r, 1).setFontSize(14).setFontWeight('bold');
  dashboard.getRange(r, 3).setValue('最終更新: ' +
    Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
  r += 2;

  // 今日の予約
  dashboard.getRange(r, 1).setValue('■ 今日の予約一覧（' + today + '）').setFontWeight('bold');
  r++;
  if (todayReservations.length === 0) {
    dashboard.getRange(r, 1).setValue('  予約なし');
    r++;
  } else {
    dashboard.getRange(r, 1).setValue('時間');
    dashboard.getRange(r, 2).setValue('氏名');
    dashboard.getRange(r, 3).setValue('メニュー');
    r++;
    for (var t = 0; t < todayReservations.length; t++) {
      dashboard.getRange(r, 1).setValue(todayReservations[t].time);
      dashboard.getRange(r, 2).setValue(todayReservations[t].name);
      dashboard.getRange(r, 3).setValue(todayReservations[t].menu);
      r++;
    }
  }
  r++;

  // 今週・今月の集計
  dashboard.getRange(r, 1).setValue('■ 今週の予約件数').setFontWeight('bold');
  dashboard.getRange(r, 2).setValue(weekCount + ' 件');
  r++;
  dashboard.getRange(r, 1).setValue('■ 今週の売上合計').setFontWeight('bold');
  dashboard.getRange(r, 2).setValue('¥' + weekSales.toLocaleString());
  r++;
  dashboard.getRange(r, 1).setValue('■ 今月の予約件数').setFontWeight('bold');
  dashboard.getRange(r, 2).setValue(monthCount + ' 件');
  r++;
  dashboard.getRange(r, 1).setValue('■ 今月の売上合計').setFontWeight('bold');
  dashboard.getRange(r, 2).setValue('¥' + monthSales.toLocaleString());
  r += 2;

  // ステータス別件数
  dashboard.getRange(r, 1).setValue('■ ステータス別件数').setFontWeight('bold');
  r++;
  dashboard.getRange(r, 1).setValue('確定');
  dashboard.getRange(r, 2).setValue(statusCount['確定'] + ' 件');
  r++;
  dashboard.getRange(r, 1).setValue('キャンセル');
  dashboard.getRange(r, 2).setValue(statusCount['キャンセル'] + ' 件');
  r++;
  dashboard.getRange(r, 1).setValue('変更');
  dashboard.getRange(r, 2).setValue(statusCount['変更'] + ' 件');
  r += 2;

  // 直近7日間の予約件数（グラフ用データ）
  dashboard.getRange(r, 1).setValue('■ 直近7日間の予約件数').setFontWeight('bold');
  r++;
  var chartStartRow = r;
  var keys = Object.keys(last7days);
  for (var k = 0; k < keys.length; k++) {
    dashboard.getRange(r, 1).setValue(keys[k]);
    dashboard.getRange(r, 2).setValue(last7days[keys[k]]);
    r++;
  }

  // グラフを作成
  try {
    // 既存のグラフを削除
    var charts = dashboard.getCharts();
    for (var c = 0; c < charts.length; c++) {
      dashboard.removeChart(charts[c]);
    }

    var chartRange = dashboard.getRange(chartStartRow, 1, keys.length, 2);
    var chart = dashboard.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(chartRange)
      .setPosition(chartStartRow, 4, 0, 0)
      .setOption('title', '直近7日間の予約件数')
      .setOption('width', 400)
      .setOption('height', 250)
      .build();

    dashboard.insertChart(chart);
  } catch (e) {
    Logger.log('グラフ作成エラー: ' + e.message);
  }

  r += 2;

  // 次の予約
  dashboard.getRange(r, 1).setValue('■ 次の予約').setFontWeight('bold');
  r++;
  if (nextReservation) {
    dashboard.getRange(r, 1).setValue(
      nextReservation.date + ' ' + nextReservation.time + ' ' +
      nextReservation.name + ' / ' + nextReservation.menu
    );
    // カウントダウン
    var diffMs = nextReservationTime - now;
    var diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    var diffMinutes = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
    r++;
    dashboard.getRange(r, 1).setValue('→ あと ' + diffHours + '時間' + diffMinutes + '分');
  } else {
    dashboard.getRange(r, 1).setValue('今後の予約なし');
  }
}

// ============================================================
// 内部ヘルパー関数
// ============================================================

/**
 * 予約IDを自動採番する（内部関数）
 * 形式: RSV-YYYYMMDD-XXX
 * @param {string} dateStr - 予約日（yyyy-MM-dd）
 * @return {string} 予約ID
 */
function generateReservationId_(dateStr) {
  var datePart = dateStr.replace(/-/g, '');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  var data = sheet.getDataRange().getValues();

  // 同日の予約数をカウントして連番を決定
  var count = 0;
  var prefix = 'RSV-' + datePart + '-';

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().indexOf(prefix) === 0) {
      var num = parseInt(data[i][0].toString().split('-')[2], 10);
      if (num > count) count = num;
    }
  }

  count++;
  var seqStr = ('000' + count).slice(-3);

  return prefix + seqStr;
}
