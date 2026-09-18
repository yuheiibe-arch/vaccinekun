/**
 * UI用：直近7日間の未充足拠点リストを取得する関数
 */
function getUnfulfilledLocationsForUI() {
  try {
    const sheetName = '医師未充足拠点';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const locations = [];

    // ヘッダーから列のインデックスを取得
    const dateIdx = headers.indexOf('対象日');
    const nameIdx = headers.indexOf('拠点名');
    const reservationsIdx = headers.indexOf('ワクチン予約数');
    const totalIdx = headers.indexOf('ワクチン予約枠数');
    const rateIdx = headers.indexOf('充足率');
    const morningIdx = headers.indexOf('午前医師');
    const afternoonIdx = headers.indexOf('午後医師');
    const nightIdx = headers.indexOf('夜間医師');
    
    // 今日から7日後までの日付を計算
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const sevenDaysLater = new Date(today);
    sevenDaysLater.setDate(today.getDate() + 7);

    // 2行目からデータを走査
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const targetDate = new Date(row[dateIdx]);
      
      // 7日以内のデータのみを対象
      if (targetDate >= today && targetDate < sevenDaysLater) {
        locations.push({
          id: `loc${i}`, // 一意のIDとして行番号を利用
          date: targetDate.toLocaleDateString('ja-JP'), // 'YYYY/MM/DD' 形式
          name: row[nameIdx],
          reservations: row[reservationsIdx],
          total: row[totalIdx],
          // 充足率をパーセント表示（例: 0.5 -> 50）に変換
          rate: Math.round((row[rateIdx] || 0) * 100),
          // 実数は午前・午後・夜間の医師数の合計
          actual: (row[morningIdx] || 0) + (row[afternoonIdx] || 0) + (row[nightIdx] || 0)
        });
      }
    }
    return locations;
  } catch(e) {
    // エラーが発生した場合はUI側にエラー情報を返す
    return { error: e.message };
  }
}

/**
 * UI用：指定された種類のアラートテンプレート情報を取得する関数
 */
function getTemplateDetails(alertType) {
  try {
    const sheetName = 'テンプレ';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }
    const data = sheet.getDataRange().getValues();

    // 2行目からデータを走査し、アラートタイプが一致する行を探す
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === alertType) { // A列がアラートタイプ
        return {
          template1: row[2],        // C列
          approachOptions: row[4],  // E列
          riskOptions: row[5]       // F列
        };
      }
    }
    // 見つからなかった場合
    throw new Error(`テンプレート「${alertType}」が見つかりません。`);
  } catch(e) {
    return { error: e.message };
  }
}