/**
 * Dailyアラートを生成し、ChatworkとSlackへ自動投稿する関数
 * この関数を時間主導型トリガー（例: 毎日午後5時〜6時）に設定してください。
 */
function triggerDailyAlertFromUI() {
  // ▼▼▼ 二重・多重実行防止ロック（ここから） ▼▼▼
  const lock = LockService.getScriptLock();
  // 同時に複数のトリガーが走った場合、2つ目以降はここで弾く（最大5秒待機）
  if (!lock.tryLock(5000)) {
    Logger.log('他のプロセスが実行中のためスキップしました。');
    return { success: false, message: '現在他のプロセスが実行中です。重複送信を防ぐため処理を中断しました。' };
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const todayStr = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
    const lastRunDate = props.getProperty('LAST_DAILY_ALERT_DATE');

    // 今日すでに1回送信済みなら、2回目以降は絶対に送信しない
    if (lastRunDate === todayStr) {
      Logger.log(`本日（${todayStr}）は既にDailyアラートを送信済みのため、多重実行を防止しました。`);
      return { success: false, message: `本日（${todayStr}）は既にDailyアラートを送信済みです。` };
    }
    // ▲▲▲ 二重・多重実行防止ロック（ここまで） ▲▲▲

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
    let locationListStr = ''; // Slack用
    const today = new Date();
    const postTimeStr = Utilities.formatDate(today, 'JST', 'HH:mm');
    const postDateStr = `${today.getMonth() + 1}月${today.getDate()}日（${getJapaneseDay(today)}）医師不在報告 ${postTimeStr}`;

    const startPeriod = Utilities.formatDate(tomorrow, 'JST', 'MM/dd') + `（${getJapaneseDay(tomorrow)}）`;
    const endPeriod = Utilities.formatDate(fiveDaysLater, 'JST', 'MM/dd') + `（${getJapaneseDay(fiveDaysLater)}）`;
    const periodStr = `${startPeriod}～${endPeriod}`;

    if (targetRows.length > 0) {
      messageBody = alertTemplate;
      locationListStr = targetRows.map(row => {
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
      
      // 不在がある場合、Chatworkには詳細リンクを追記する
      const detailLinkText = '5日以降の不在状況詳細はこちらをご確認ください。\nhttps://docs.google.com/spreadsheets/d/1BobYzsY2ApVTCP07qkJFOc6ZQkxq1bJgXl2endV5vHg/edit?gid=59569568#gid=59569568';
      messageBody = messageBody.replace(/([hr])[\s\S]*(\[\/info\])/, `$1\n${locationListStr}\n\n${detailLinkText}\n$2`);

    } else {
      messageBody = noAlertTemplate;
      messageBody = messageBody.replace(/◯月◯日\s*（.）医師不在報告/, postDateStr);
      messageBody = messageBody.replace(/〇〇\/〇〇～〇〇\/〇〇/, periodStr);
    }

    // ★ 送信直前に本日の実行記録を保存（エラー時はここに来ないため再実行可能）
    props.setProperty('LAST_DAILY_ALERT_DATE', todayStr);

    // --- 1. Chatworkへの投稿 ---
    // ★ 修正: 不在が1件以上ある場合のみ [toall] を付与する
    const finalMessage = (targetRows.length > 0) ? '[toall]\n' + messageBody : messageBody;
    
    const roomInfo = getRoomInfo_v2('【緊急】予約振替対策チーム[DS×CL×CS]');
    if (roomInfo.error) throw new Error(roomInfo.error);
    
    postAlertToChatwork_v2(roomInfo.id, finalMessage, {}); 

    // --- 2. Slackへの投稿 ---
    postAlertToSlack(targetRows.length, postDateStr, periodStr, locationListStr);

    return { success: true, message: 'Dailyアラートの投稿が完了しました。' };

  } catch(e) {
    Logger.log(`Dailyアラート生成・投稿エラー: ${e.message}\n${e.stack}`);
    throw new Error(`Dailyアラートの作成に失敗しました: ${e.message}`);
  } finally {
    // 処理が終わったら必ずロックを解放する
    lock.releaseLock();
  }
}

/**
 * ==========================================
 * Slack通知用関数 (Block Kit対応・画像フォーマット再現版)
 * ==========================================
 */
function postAlertToSlack(targetCount, postDateStr, periodStr, locationListStr) {
  const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!SLACK_WEBHOOK_URL) {
    Logger.log("❌ Slack Webhook URL が設定されていません。");
    return;
  }

  const sheetUrl = 'https://docs.google.com/spreadsheets/d/1BobYzsY2ApVTCP07qkJFOc6ZQkxq1bJgXl2endV5vHg/edit?gid=59569568#gid=59569568';
  
  // スマホ通知などに表示されるプレビューテキスト (不在0件時はメンションなし)
  let previewText = targetCount > 0 
    ? `<!channel> 🚨 ワクチン医師不在報告: ${targetCount}件の不在があります` 
    : `✅ ワクチン医師不在報告: 対象期間に医師不在はありません。`;

  let blocks = [];

  if (targetCount > 0) {
    // 不在拠点がある場合のレイアウト (<!channel> 付き)
    blocks = [
      {
        "type": "header",
        "text": {
          "type": "plain_text",
          "text": "🚨 ワクチン医師不在報告",
          "emoji": true
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": `<!channel>\nDaily checkが完了しました。\n対象期間（${periodStr}）において、*${targetCount}件* の医師不在拠点があります。`
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": `*【不在拠点一覧】*\n${locationListStr}`
        }
      },
      {
        "type": "actions",
        "elements": [
          {
            "type": "button",
            "text": {
              "type": "plain_text",
              "text": "📊 不在状況詳細を開く (スプレッドシート)",
              "emoji": true
            },
            "url": sheetUrl,
            "style": "primary"
          }
        ]
      },
      {
        "type": "divider"
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "💡 *【担当者へのお願い】*\nリンク先のスプレッドシートにて、5日以降の不在状況の詳細や最新状況をご確認ください。"
        }
      }
    ];
  } else {
    // 不在が0件の場合のシンプルなレイアウト (メンションなし)
    blocks = [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": `対象期間（${postDateStr}）\n\n対象期間に医師不在はありません。`
        }
      }
    ];
  }

  const payload = {
    "channel": "C0C3KKZ92NL",
    "text": previewText,
    "blocks": blocks
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    Logger.log("✅ Slack通知の送信に成功しました。");
  } catch (e) {
    Logger.log("❌ Slack送信エラー: " + e.message);
  }
}