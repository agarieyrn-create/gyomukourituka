// ============================================================
// mail.gs - メール送信処理
// GAS 予約管理ツール
// ============================================================

/**
 * 予約確認メールを送信する
 * @param {Object} reservation - 予約データ
 */
function sendConfirmationMail(reservation) {
  if (!reservation.email) return;

  var settings = getSettings();
  var shopName = settings['店舗名'] || '当店';

  var subject = '【' + shopName + '】ご予約を承りました';

  var body = createMailHtml_({
    title: 'ご予約確認',
    shopName: shopName,
    greeting: reservation.name + ' 様',
    message: 'ご予約を承りました。以下の内容をご確認ください。',
    details: [
      { label: '予約ID', value: reservation.reservationId },
      { label: '予約日', value: reservation.date },
      { label: '予約時間', value: reservation.time },
      { label: 'メニュー', value: reservation.menuName },
      { label: '料金', value: '¥' + reservation.price.toLocaleString() }
    ],
    footer: '※キャンセル・変更の際は予約IDが必要です。このメールを大切に保管してください。'
  });

  try {
    GmailApp.sendEmail(reservation.email, subject, '', {
      htmlBody: body,
      name: shopName
    });
  } catch (e) {
    Logger.log('確認メール送信エラー: ' + e.message);
  }
}

/**
 * 変更完了メールを送信する
 * @param {Object} reservation - 変更後の予約データ
 */
function sendChangeMail(reservation) {
  if (!reservation.email) return;

  var settings = getSettings();
  var shopName = settings['店舗名'] || '当店';

  var subject = '【' + shopName + '】ご予約を変更しました';

  var body = createMailHtml_({
    title: '予約変更完了',
    shopName: shopName,
    greeting: reservation.name + ' 様',
    message: 'ご予約の変更が完了しました。変更後の内容は以下の通りです。',
    details: [
      { label: '予約ID', value: reservation.reservationId },
      { label: '変更後の予約日', value: reservation.date },
      { label: '変更後の予約時間', value: reservation.time },
      { label: 'メニュー', value: reservation.menuName },
      { label: '料金', value: '¥' + reservation.price.toLocaleString() }
    ],
    footer: '※再度変更・キャンセルの際は予約IDをご用意ください。'
  });

  try {
    GmailApp.sendEmail(reservation.email, subject, '', {
      htmlBody: body,
      name: shopName
    });
  } catch (e) {
    Logger.log('変更メール送信エラー: ' + e.message);
  }
}

/**
 * キャンセル完了メールを送信する
 * @param {Object} reservation - 予約データ
 */
function sendCancelMail(reservation) {
  if (!reservation.email) return;

  var settings = getSettings();
  var shopName = settings['店舗名'] || '当店';

  var subject = '【' + shopName + '】ご予約をキャンセルしました';

  var body = createMailHtml_({
    title: 'キャンセル完了',
    shopName: shopName,
    greeting: reservation.name + ' 様',
    message: 'ご予約のキャンセルが完了しました。',
    details: [
      { label: '予約ID', value: reservation.reservationId },
      { label: '予約日', value: reservation.date },
      { label: '予約時間', value: reservation.time },
      { label: 'メニュー', value: reservation.menuName }
    ],
    footer: 'またのご利用をお待ちしております。'
  });

  try {
    GmailApp.sendEmail(reservation.email, subject, '', {
      htmlBody: body,
      name: shopName
    });
  } catch (e) {
    Logger.log('キャンセルメール送信エラー: ' + e.message);
  }
}

/**
 * リマインダーメールを送信する
 * @param {Object} reservation - 予約データ
 */
function sendReminderMail(reservation) {
  if (!reservation.email) return;

  var settings = getSettings();
  var shopName = settings['店舗名'] || '当店';

  var subject = '【' + shopName + '】ご予約のリマインダー';

  var body = createMailHtml_({
    title: 'ご予約リマインダー',
    shopName: shopName,
    greeting: reservation.name + ' 様',
    message: 'ご予約日が近づいてまいりましたのでお知らせいたします。',
    details: [
      { label: '予約ID', value: reservation.reservationId },
      { label: '予約日', value: reservation.date },
      { label: '予約時間', value: reservation.time },
      { label: 'メニュー', value: reservation.menuName }
    ],
    footer: 'ご都合が悪くなった場合はお早めにご連絡ください。'
  });

  try {
    GmailApp.sendEmail(reservation.email, subject, '', {
      htmlBody: body,
      name: shopName
    });
  } catch (e) {
    Logger.log('リマインダーメール送信エラー: ' + e.message);
  }
}

/**
 * 管理者通知メールを送信する
 * @param {Object} reservation - 予約データ
 * @param {string} type - 通知種別（予約/変更/キャンセル）
 */
function sendAdminNotification(reservation, type) {
  var settings = getSettings();
  var adminEmail = settings['管理者メールアドレス'];
  if (!adminEmail) return;

  var shopName = settings['店舗名'] || '当店';
  var subject = '【' + shopName + '管理者通知】' + type + 'がありました';

  var details = [
    { label: '予約ID', value: reservation.reservationId },
    { label: '氏名', value: reservation.name },
    { label: '予約日', value: reservation.date },
    { label: '予約時間', value: reservation.time },
    { label: 'メニュー', value: reservation.menuName }
  ];

  if (reservation.phone) {
    details.push({ label: '電話番号', value: reservation.phone });
  }
  if (reservation.email) {
    details.push({ label: 'メールアドレス', value: reservation.email });
  }

  var body = createMailHtml_({
    title: '管理者通知 - ' + type,
    shopName: shopName,
    greeting: '管理者様',
    message: '以下の' + type + 'がありました。',
    details: details,
    footer: ''
  });

  try {
    GmailApp.sendEmail(adminEmail, subject, '', {
      htmlBody: body,
      name: shopName + ' 予約システム'
    });
  } catch (e) {
    Logger.log('管理者通知メール送信エラー: ' + e.message);
  }
}

// ============================================================
// メールテンプレート（内部関数）
// ============================================================

/**
 * HTML形式のメール本文を生成する（内部関数）
 * @param {Object} params - メールパラメータ
 * @return {string} HTML文字列
 */
function createMailHtml_(params) {
  var detailsHtml = '';
  for (var i = 0; i < params.details.length; i++) {
    var d = params.details[i];
    detailsHtml +=
      '<tr>' +
      '<td style="padding:8px 12px;background:#f8f9fa;border:1px solid #e9ecef;font-weight:bold;width:140px;color:#495057;">' +
      d.label + '</td>' +
      '<td style="padding:8px 12px;border:1px solid #e9ecef;color:#212529;">' +
      d.value + '</td>' +
      '</tr>';
  }

  return '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8"></head>' +
    '<body style="margin:0;padding:0;background:#f5f5f5;font-family:\'Helvetica Neue\',Arial,\'Hiragino Kaku Gothic ProN\',\'Hiragino Sans\',Meiryo,sans-serif;">' +
    '<div style="max-width:600px;margin:20px auto;background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">' +

    // ヘッダー
    '<div style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:24px;text-align:center;">' +
    '<h1 style="color:#ffffff;margin:0;font-size:20px;letter-spacing:1px;">' + params.shopName + '</h1>' +
    '</div>' +

    // タイトル
    '<div style="padding:24px 24px 0;">' +
    '<h2 style="color:#333;font-size:18px;margin:0 0 16px;padding-bottom:12px;border-bottom:2px solid #667eea;">' +
    params.title + '</h2>' +
    '</div>' +

    // 本文
    '<div style="padding:0 24px;">' +
    '<p style="color:#555;line-height:1.8;margin:16px 0;">' + params.greeting + '</p>' +
    '<p style="color:#555;line-height:1.8;margin:0 0 20px;">' + params.message + '</p>' +

    // 詳細テーブル
    '<table style="width:100%;border-collapse:collapse;margin:0 0 20px;">' +
    detailsHtml +
    '</table>' +

    // フッターメッセージ
    (params.footer ?
      '<p style="color:#888;font-size:13px;line-height:1.6;margin:20px 0;padding:12px;background:#fff8e1;border-radius:4px;border-left:4px solid #ffc107;">' +
      params.footer + '</p>' : '') +
    '</div>' +

    // フッター
    '<div style="padding:16px 24px;background:#f8f9fa;text-align:center;border-top:1px solid #e9ecef;">' +
    '<p style="color:#999;font-size:12px;margin:0;">このメールは ' + params.shopName + ' 予約システムから自動送信されています。</p>' +
    '</div>' +

    '</div></body></html>';
}
