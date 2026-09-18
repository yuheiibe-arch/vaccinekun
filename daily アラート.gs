/**
 * Dailyアラートを生成し、Chatworkへ自動投稿する関数
 * この関数を時間主導型トリガー（例: 毎日午後5時〜6時）に設定してください。
 */
function triggerDailyAlertFromUI() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const unfulfilledSheet = ss.getSheetByName('医師未充足拠点');
    const templateSheet = ss.getSheetByName('テンプレ');

    if (!unfulfilledSheet) throw new Error('シート「医師未充足拠点」が見つかりません。');
    if (!templateSheet) throw new Error('シート「テンプレ」が見つかりません。');

    const allData = unfulfilledSheet.getDataRange().getValues();
    const headers = allData[0];
    
    const requiredHeaders = ['対象日', '拠点名', '午前医師', '午後医師', '夜間医師', 'ワクチン予約数', 'ワクチン予約枠数', '充足率'];
    const headerIndices = {};
    requiredHeaders.forEach(header => {
      const index = headers.indexOf(header);
      if (index === -1) {
        throw new Error(`「医師未充足拠点」シートに必須ヘッダー「${header}」が見つかりません。`);
      }
      headerIndices[header] = index;
    });

    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    tomorrow.setHours(0, 0, 0, 0);

    const fiveDaysLater = new Date();
    fiveDaysLater.setDate(fiveDaysLater.getDate() + 5);
    fiveDaysLater.setHours(23, 59, 59, 999);

    const targetRows = allData.slice(1).filter(row => {
      if (!row[headerIndices['対象日']]) return false;
      const targetDate = new Date(row[headerIndices['対象日']]);
      return targetDate >= tomorrow && targetDate <= fiveDaysLater;
    });

    const templateData = templateSheet.getDataRange().getValues();
    let alertTemplate = '', noAlertTemplate = '';
    for (let i = 1; i < templateData.length; i++) {
      if (templateData[i][0] === 'Daily Alert') {
        alertTemplate = templateData[i][2];
        noAlertTemplate = templateData[i][3];
        break;
      }
    }
    if (!alertTemplate || !noAlertTemplate) throw new Error('「テンプレ」シートのA列に項目名 "Daily Alert" が見つかりません。');

    let messageBody = '';
    const today = new Date();
    const postTimeStr = Utilities.formatDate(today, 'JST', 'HH:mm');
    const postDateStr = `${today.getMonth() + 1}月${today.getDate()}日（${getJapaneseDay(today)}）医師不在報告　${postTimeStr}`;

    const startPeriod = Utilities.formatDate(tomorrow, 'JST', 'MM/dd') + `（${getJapaneseDay(tomorrow)}）`;
    const endPeriod = Utilities.formatDate(fiveDaysLater, 'JST', 'MM/dd') + `（${getJapaneseDay(fiveDaysLater)}）`;
    const periodStr = `${startPeriod}～${endPeriod}`;

    if (targetRows.length > 0) {
      messageBody = alertTemplate;
      const locationListStr = targetRows.map(row => {
        const targetDate = new Date(row[headerIndices['対象日']]);
        const dateHeader = Utilities.formatDate(targetDate, 'JST', 'MM/dd') + `（${getJapaneseDay(targetDate)}）`;

        const unfulfilledSlots = [];
        if (row[headerIndices['午前医師']] == 0) unfulfilledSlots.push('午前');
        if (row[headerIndices['午後医師']] == 0) unfulfilledSlots.push('午後');
        if (row[headerIndices['夜間医師']] == 0) unfulfilledSlots.push('夜間');
        const timeStr = unfulfilledSlots.join('・');

        const name = row[headerIndices['拠点名']];
        const reserv = row[headerIndices['ワクチン予約数']];
        const total = row[headerIndices['ワクチン予約枠数']];
        const rate = Math.round((row[headerIndices['充足率']] || 0) * 100);

        return `${dateHeader}\n【${name}】 ${timeStr}\nワクチン予約数：${reserv}件\n充足率：${rate}％（${reserv}/${total}）`;
      }).join('\n\n');
      
      messageBody = messageBody.replace(/◯月◯日\s*（.）医師不在報告/, postDateStr);
      messageBody = messageBody.replace(/〇〇\/〇〇～〇〇\/〇〇/, periodStr);
      messageBody = messageBody.replace(/(\[hr\])[\s\S]*(\[\/info\])/, `$1\n${locationListStr}\n$2`);

    } else {
      messageBody = noAlertTemplate;
      messageBody = messageBody.replace(/◯月◯日\s*（.）医師不在報告/, postDateStr);
      messageBody = messageBody.replace(/〇〇\/〇〇～〇〇\/〇〇/, periodStr);
    }

    const finalMessage = '[toall]\n' + messageBody;
    const roomInfo = getRoomInfo_v2('【緊急】予約振替対策チーム[DS×CL×CS]');
    if (roomInfo.error) throw new Error(roomInfo.error);
    
    postAlertToChatwork_v2(roomInfo.id, finalMessage, {}); 

    // トリガー実行の場合、この戻り値はログに記録されるだけです
    return { success: true, message: 'Dailyアラートの投稿が完了しました。' };

  } catch(e) {
    Logger.log(`Dailyアラート生成・投稿エラー: ${e.message}\n${e.stack}`);
    throw new Error(`Dailyアラートの作成に失敗しました: ${e.message}`);
  }
}

/**
 * 日本語の曜日を取得するヘルパー関数
 */
function getJapaneseDay(date) {
  return ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
}