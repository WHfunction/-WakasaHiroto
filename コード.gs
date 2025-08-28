/**
 * Webアプリとしてアクセスされたときに呼び出されるメイン関数
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('若狭裕人マイポータル')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * HTMLファイル内で他のHTMLファイルをインクルードするための関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gmailの受信トレイから最新のメッセージを取得するサーバーサイド関数
 */
function getGmailMessages() {
  try {
    const threads = GmailApp.getInboxThreads(0, 20); 
    const emailData = [];
    threads.forEach(thread => {
      const message = thread.getMessages()[thread.getMessageCount() - 1];
      emailData.push({
        id: message.getId(),
        threadId: thread.getId(),
        subject: message.getSubject(),
        from: message.getFrom(),
        date: message.getDate().toLocaleString('ja-JP'),
        snippet: message.getPlainBody().substring(0, 100),
        isUnread: message.isUnread()
      });
    });
    return emailData;
  } catch (e) {
    console.error("Gmailの取得に失敗しました: " + e.toString());
    return { error: 'メールの取得中にエラーが発生しました。' };
  }
}

/**
 * Gmailの未読メール件数を取得する関数
 */
function getUnreadMailCount() {
  try {
    return GmailApp.getInboxUnreadCount();
  } catch(e) {
    console.error("未読件数の取得に失敗: " + e.toString());
    return 0;
  }
}

/**
 * 指定された日付に提出物があるかチェックする内部関数
 * @param {Date} targetDate チェックする日付
 * @returns {boolean} 提出物があればtrue
 */
function hasSubmissionsOn(targetDate) {
  const keywords = ["提出", "締切", "締め切り", "レポート", "課題", "Assignment"];
  const startTime = new Date(targetDate);
  startTime.setHours(0, 0, 0, 0);
  const endTime = new Date(targetDate);
  endTime.setHours(23, 59, 59, 999);

  const calendars = CalendarApp.getAllCalendars();
  for (const calendar of calendars) {
    const events = calendar.getEvents(startTime, endTime);
    for (const event of events) {
      const title = event.getTitle();
      if (keywords.some(keyword => title.toLowerCase().includes(keyword.toLowerCase()))) {
        return true;
      }
    }
  }
  return false;
}

/**
 * 今日と明日の提出物の有無をまとめて取得する
 * @returns {Object} { today: boolean, tomorrow: boolean }
 */
function getSubmissionStatus() {
  try {
    const today = new Date();
    const tomorrow = new Date();
    tomorrow.setDate(today.getDate() + 1);

    return {
      today: hasSubmissionsOn(today),
      tomorrow: hasSubmissionsOn(tomorrow)
    };
  } catch(e) {
    console.error("カレンダーのチェックに失敗: " + e.toString());
    return { today: false, tomorrow: false };
  }
}
