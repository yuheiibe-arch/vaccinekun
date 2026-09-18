/**
 * =======================================================================================
 * ワクチンくんUI - サーバー側スクリプト
 * =======================================================================================
 */

function openUIMain() {
  const htmlTemplate = HtmlService.createTemplateFromFile('UIMain');
  const html = htmlTemplate.evaluate()
    .setWidth(1100)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'ワクチンくん 投稿ツール');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ワクチンくん起動')
    .addItem('ワクチンくん発動', 'openUIMain')
    .addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 日本語の曜日を取得するヘルパー関数
 */
function getJapaneseDay(date) {
  return ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
}

/**
 * UI用：Dailyアラートのプレビュー文を生成して返す関数
 */
function generateDailyAlertPreview() {
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
    const postDateStr = `${today.getMonth() + 1}月${today.getDate()}日（${getJapaneseDay(today)}）医師不在報告 ${postTimeStr}`;
    
    const startPeriod = Utilities.formatDate(tomorrow, 'JST', 'MM/dd') + `（${getJapaneseDay(tomorrow)}）`;
    const endPeriod = Utilities.formatDate(fiveDaysLater, 'JST', 'MM/dd') + `（${getJapaneseDay(fiveDaysLater)}）`;
    const periodStr = `${startPeriod}～${endPeriod}`;

    const detailLinkText = '5日以降の不在状況詳細はこちらをご確認ください。\nhttps://docs.google.com/spreadsheets/d/1BobYzsY2ApVTCP07qkJFOc6ZQkxq1bJgXl2endV5vHg/edit?gid=59569568#gid=59569568';

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
      messageBody = messageBody.replace(/(\[hr\])[\s\S]*(\[\/info\])/, `$1\n${locationListStr}\n\n${detailLinkText}\n$2`);

    } else {
      messageBody = noAlertTemplate;
      messageBody = messageBody.replace(/◯月◯日\s*（.）医師不在報告/, postDateStr);
      messageBody = messageBody.replace(/〇〇\/〇〇～〇〇\/〇〇/, periodStr);
      messageBody = messageBody.replace(/(\[hr\])/, `$1\n${detailLinkText}`);
    }
    
    return messageBody;

  } catch(e) {
    Logger.log(`Dailyアラートのプレビュー生成エラー: ${e.message}\n${e.stack}`);
    return { error: e.message };
  }
}

// --- これ以降の関数は変更ありません ---

function getUnfulfilledLocationsForUI_v2() { 
  try {
    const sheetName = '医師未充足拠点';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const locations = [];
    const dateIdx = headers.indexOf('対象日'), nameIdx = headers.indexOf('拠点名'), reservationsIdx = headers.indexOf('ワクチン予約数');
    const totalIdx = headers.indexOf('ワクチン予約枠数'), rateIdx = headers.indexOf('充足率'), morningIdx = headers.indexOf('午前医師');
    const afternoonIdx = headers.indexOf('午後医師'), nightIdx = headers.indexOf('夜間医師');
    
    if ([dateIdx, nameIdx, morningIdx, afternoonIdx, nightIdx].includes(-1)) throw new Error('「医師未充足拠点」シートに必要なヘッダーが見つかりません。');
    
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const sevenDaysLater = new Date(today); sevenDaysLater.setDate(today.getDate() + 7);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[dateIdx]) continue;
      const targetDate = new Date(row[dateIdx]);
      
      if (targetDate >= today && targetDate < sevenDaysLater) {
        const unfulfilledSlots = [];
        if (row[morningIdx] == 0) unfulfilledSlots.push('午前');
        if (row[afternoonIdx] == 0) unfulfilledSlots.push('午後');
        if (row[nightIdx] == 0) unfulfilledSlots.push('夜間');

        if (unfulfilledSlots.length === 0) continue;
        const unfulfilledTime = unfulfilledSlots.join('・');
        
        locations.push({
          id: `loc${i}`, date: targetDate.toLocaleDateString('ja-JP'), name: row[nameIdx],
          reservations: row[reservationsIdx], total: row[totalIdx], rate: Math.round((row[rateIdx] || 0) * 100),
          actual: (Number(row[morningIdx]) || 0) + (Number(row[afternoonIdx]) || 0) + (Number(row[nightIdx]) || 0),
          unfulfilledTime: unfulfilledTime
        });
      }
    }
    return locations;
  } catch(e) { return { error: e.message }; }
}

function getEditorData_v2(alertType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('テンプレ');
    if (!templateSheet) throw new Error('シート「テンプレ」が見つかりません。');
    const templateData = templateSheet.getDataRange().getValues();
    const alertTypeMap = { 'unfulfilled': '未充足危険アラート', 'transfer': '振替依頼アラート' };
    const targetItemName = alertTypeMap[alertType];
    if (!targetItemName) throw new Error(`不正なアラートタイプです: ${alertType}`);
    let templateInfo = null;
    for (let i = 1; i < templateData.length; i++) {
      const row = templateData[i];
      if (row[0] === targetItemName) {
        templateInfo = { to: row[1], template1: row[2], approachOptions: row[4], riskOptions: row[5], reasons: row[6] };
        break;
      }
    }
    if (!templateInfo) throw new Error(`「テンプレ」シートに項目「${targetItemName}」が見つかりません。`);

    const mentionSheet = ss.getSheetByName('メンションリスト');
    if (!mentionSheet) throw new Error('シート「メンションリスト」が見つかりません。');
    const mentionData = mentionSheet.getDataRange().getValues();
    const mentionLists = { ds: [], ccs: [], others: [] };
    for (let i = 1; i < mentionData.length; i++) {
      if (mentionData[i][0]) mentionLists.ds.push(mentionData[i][0]);
      if (mentionData[i][1]) mentionLists.ccs.push(mentionData[i][1]);
      if (mentionData[i][2]) mentionLists.others.push(mentionData[i][2]);
    }
    
    const fiscalYearAlertCount = getFiscalYearAlertCount(ss);
    return { template: templateInfo, mentions: mentionLists, fiscalYearAlertCount: fiscalYearAlertCount };

  } catch(e) { return { error: e.message }; }
}

function getRoomInfo_v2(roomName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('データ');
    if (!sheet) throw new Error('シート「データ」が見つかりません。');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === roomName) return { name: data[i][0], id: data[i][1] };
    }
    throw new Error(`シート「データ」に「${roomName}」の情報が見つかりません。`);
  } catch(e) { return { error: e.message }; }
}

function postAlertToChatwork_v2(roomId, message, historyData) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
    if (!token) throw new Error('スクリプトプロパティに「CHATWORK_API_TOKEN」が設定されていません。');
    
    const endpoint = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
    const payload = { 'body': message };
    const options = { 'method': 'post', 'headers': { 'X-ChatWorkToken': token }, 'payload': payload };
    
    const response = UrlFetchApp.fetch(endpoint, options);
    
    if (historyData && historyData.alertType === 'transfer') {
      logAlertHistory(historyData);
    }
    
    Logger.log(`投稿成功: ${response.getContentText()}`);
    return { success: true, message: '投稿が完了しました。' };
  } catch(e) {
    Logger.log(`投稿失敗: ${e.message}`);
    return { success: false, message: `投稿に失敗しました: ${e.message}` };
  }
}

function logAlertHistory(historyData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('通知履歴');
  if (!historySheet) return;
  const postDate = new Date();
  const fiscalYear = getFiscalYear(postDate);
  const timestamp = Utilities.formatDate(postDate, 'JST', 'yyyy-MM-dd HH:mm:ss');
  historyData.locations.forEach(location => {
    historySheet.appendRow([ timestamp, fiscalYear, location.name, location.unfulfilledTime, postDate, historyData.reason, historyData.fiscalYearAlertCount ]);
  });
}

function getFiscalYearAlertCount(spreadsheet) {
  const historySheet = spreadsheet.getSheetByName('通知履歴');
  if (!historySheet || historySheet.getLastRow() < 2) return 1;
  const data = historySheet.getRange(2, 2, historySheet.getLastRow() - 1, 1).getValues();
  const currentFiscalYear = getFiscalYear(new Date());
  let count = 0;
  for (let i = 0; i < data.length; i++) { if (data[i][0] == currentFiscalYear) count++; }
  return count + 1;
}

function getFiscalYear(date) {
  const year = date.getFullYear(); const month = date.getMonth() + 1;
  return (month >= 4) ? year : year - 1;
}