/**
 * ==========================================
 * Slack通知用関数 (Block Kit対応・ワクチンくん版)
 * ==========================================
 */
function postAlertToSlack(targetCount, periodStr, locationListStr) {
  // Webhook URLをスクリプトプロパティから取得
  const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!SLACK_WEBHOOK_URL) {
    Logger.log("❌ Slack Webhook URL が設定されていません。");
    return;
  }

  const sheetUrl = 'https://docs.google.com/spreadsheets/d/1BobYzsY2ApVTCP07qkJFOc6ZQkxq1bJgXl2endV5vHg/edit?gid=59569568#gid=59569568';
  
  // スマホ通知などに表示されるプレビューテキスト
  let previewText = targetCount > 0 
    ? `@dspart @dsshift 🚨 ワクチンくん: ${periodStr} の期間に ${targetCount}件 の医師不在があります` 
    : `✅ ワクチンくん: ${periodStr} の期間に医師の不在はありません`;

  let blocks = [];

  if (targetCount > 0) {
    // 不在拠点がある場合のレイアウト
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
          "text": `@dspart @dsshift\nDaily checkが完了しました。\n対象期間（${periodStr}）において、*${targetCount}件* の医師不在拠点があります。`
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
    // 不在が0件の場合のシンプルなレイアウト（定型文のみ）
    blocks = [
      {
        "type": "header",
        "text": {
          "type": "plain_text",
          "text": "✅ ワクチン医師不在報告",
          "emoji": true
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": `@dspart @dsshift\n対象期間（${periodStr}）\n\n\`対象期間に医師不在はありません。\``
        }
      }
    ];
  }

  const payload = {
    "channel": "C0C3KKZ92NL", // 指定されたSlackチャンネルID
    "text": previewText,
    "blocks": blocks
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  // Slackへ送信
  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    Logger.log("✅ Slack通知の送信に成功しました。");
  } catch (e) {
    Logger.log("❌ Slack送信エラー: " + e.message);
  }
}