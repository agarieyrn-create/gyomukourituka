// ============================================================
// Code.gs - メイン・ルーティング
// GAS 予約管理ツール
// ============================================================

/**
 * GETリクエスト処理
 * pageパラメータに応じて表示するHTMLを切り替える
 * @param {Object} e - リクエストパラメータ
 * @return {HtmlOutput} HTMLページ
 */
function doGet(e) {
  var page = e.parameter.page || 'index';
  var settings = getSettings();
  var shopName = settings['店舗名'] || '予約システム';

  switch (page) {
    case 'cancel':
      return HtmlService.createHtmlOutputFromFile('cancel')
        .setTitle(shopName + ' - キャンセル')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    case 'change':
      return HtmlService.createHtmlOutputFromFile('change')
        .setTitle(shopName + ' - 予約変更')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    default:
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle(shopName + ' - ご予約')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * POSTリクエスト処理
 * actionパラメータに応じて各処理を呼び出す
 * @param {Object} e - リクエストパラメータ
 * @return {TextOutput} JSON形式のレスポンス
 */
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    var result;

    switch (action) {
      case 'getAvailableSlots':
        // 指定日の空き枠を返す
        result = getDaySlots(params.date);
        break;

      case 'getCalendarData':
        // 月ごとの空き状況を返す
        result = getMonthAvailability(params.year, params.month);
        break;

      case 'submitReservation':
        // 予約を登録する
        result = handleSubmitReservation_(params);
        break;

      case 'cancelReservation':
        // キャンセルを処理する
        result = processCancellation(params.reservationId);
        break;

      case 'changeReservation':
        // 予約変更を処理する
        result = processChange(params.reservationId, params.newDate, params.newTime);
        break;

      case 'findReservation':
        // 予約IDと氏名で予約を検索する
        result = findReservation_(params.reservationId, params.name);
        break;

      case 'getMenuList':
        // メニュー一覧を取得する
        result = getMenuList();
        break;

      case 'getShopSettings':
        // 店舗設定を取得する
        result = getPublicSettings_();
        break;

      default:
        throw new Error('不明なアクションです: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: result
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('doPost エラー: ' + error.message + '\n' + error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 内部関数
// ============================================================

/**
 * 予約登録のメイン処理（内部関数）
 * LockServiceで二重予約を防止する
 * @param {Object} params - 予約パラメータ
 * @return {Object} 予約結果
 */
function handleSubmitReservation_(params) {
  // LockServiceで同時リクエストによる二重予約を防止
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 最大10秒待機
  } catch (e) {
    throw new Error('サーバーが混雑しています。しばらくしてから再度お試しください。');
  }

  try {
    // メニュー情報を取得
    var menuList = getMenuList();
    var selectedMenu = null;
    for (var i = 0; i < menuList.length; i++) {
      if (menuList[i].menuId === params.menuId) {
        selectedMenu = menuList[i];
        break;
      }
    }
    if (!selectedMenu) {
      throw new Error('選択されたメニューが見つかりません。');
    }

    // 重複チェック
    var duration = selectedMenu.duration;
    if (checkDuplicate(params.date, params.time, duration, null)) {
      throw new Error('選択された時間帯はすでに予約が入っています。');
    }

    // 予約データを作成
    var reservationData = {
      date: params.date,
      time: params.time,
      name: params.name,
      phone: params.phone,
      email: params.email,
      menuId: params.menuId,
      menuName: selectedMenu.name,
      duration: duration,
      price: selectedMenu.price,
      remarks: params.remarks || ''
    };

    // スプレッドシートに保存
    var reservationId = saveReservation(reservationData);
    reservationData.reservationId = reservationId;

    // Googleカレンダーにイベント登録
    var eventId = registerCalendarEvent(reservationData);

    // カレンダーイベントIDを予約一覧に記録
    updateEventId_(reservationId, eventId);

    // 確認メール送信
    sendConfirmationMail(reservationData);

    // 管理者通知メール送信
    sendAdminNotification(reservationData, '予約');

    // ダッシュボード更新
    updateDashboard();

    return {
      reservationId: reservationId,
      name: reservationData.name,
      date: reservationData.date,
      time: reservationData.time,
      menuName: reservationData.menuName
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * 予約検索（内部関数）
 * 予約IDと氏名で一致する予約を返す
 * @param {string} reservationId - 予約ID
 * @param {string} name - 氏名
 * @return {Object} 予約データ
 */
function findReservation_(reservationId, name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  if (!sheet) throw new Error('予約一覧シートが見つかりません。');

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] === reservationId && row[4] === name) {
      return {
        reservationId: row[0],
        receiveDate: row[1],
        date: Utilities.formatDate(new Date(row[2]), 'Asia/Tokyo', 'yyyy-MM-dd'),
        time: row[3],
        name: row[4],
        phone: row[5],
        email: row[6],
        menuName: row[7],
        price: row[8],
        status: row[9],
        remarks: row[10]
      };
    }
  }

  throw new Error('該当する予約が見つかりません。予約IDと氏名をご確認ください。');
}

/**
 * カレンダーイベントIDを予約一覧に記録する（内部関数）
 * @param {string} reservationId - 予約ID
 * @param {string} eventId - カレンダーイベントID
 */
function updateEventId_(reservationId, eventId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('予約一覧');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      sheet.getRange(i + 1, 12).setValue(eventId); // L列：カレンダーイベントID
      return;
    }
  }
}

/**
 * 公開用の設定情報を返す（内部関数）
 * 管理者メールアドレスなど非公開情報は含めない
 * @return {Object} 公開設定
 */
function getPublicSettings_() {
  var settings = getSettings();
  return {
    shopName: settings['店舗名'] || '',
    openTime: settings['営業開始時間'] || '10:00',
    closeTime: settings['営業終了時間'] || '19:00',
    interval: parseInt(settings['予約間隔（分）'] || '60', 10),
    maxBookingDays: parseInt(settings['予約受付期間（何日先まで）'] || '30', 10)
  };
}
