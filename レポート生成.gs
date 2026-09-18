/**
 * =======================================================================================
 * レポート生成ファイル
 * =======================================================================================
 * v21 修正点:
 * - 通常レポートの「充足」の判定条件を「午前・午後・夜間の全てに医師がいる場合」に限定。
 */

/**
 * 全てのレポートを更新するメインの処理フロー
 */
function updateAllReports() {
  // --- 基本設定 ---
  const masterSheetName = 'データ';
  const regularReportNumDays = 14;
  const dataFetchNumDays = 30;
  const regularReportSheetName = '予約リスト';
  const specialistReportSheetName = '専任シフト予約リスト';
  const unfulfilledSheetName = '医師未充足拠点';
  const specialistShiftSpreadsheetId = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA';
  const specialistShiftSheetName = '貼付用';

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    Logger.log('--- 全レポート更新処理を開始 ---');

    // --- STEP 1: 全レポートで共通利用するデータを準備 ---
    Logger.log('1. 各種マスタデータを準備します...');
    const regexMap = getRegexMap(masterSheetName, '正規表現');
    const latestDataMap = getLatestDataMap(masterSheetName, 'ワクチン予約数', 'データ追加', dataFetchNumDays, regexMap);
    const reservationInfo = getSheetInfo(masterSheetName, 'ワクチン予約数');
    const reservationSpreadsheet = SpreadsheetApp.openByUrl(reservationInfo.url);
    const reservationSheet = reservationSpreadsheet.getSheetByName(reservationInfo.sheetName);
    const reservationData = reservationSheet.getDataRange().getValues();
    const specialistShiftMap = createSpecialistShiftMap(specialistShiftSpreadsheetId, specialistShiftSheetName, regexMap);
    Logger.log(' -> 準備完了');

    // --- STEP 2: 通常シフトのレポートを作成 ---
    Logger.log('2. 通常シフトの「予約リスト」を作成します...');
    const regularShiftMap = getShiftMap(masterSheetName, '確定シフト', dataFetchNumDays, regexMap);
    const normalReportData = createNormalReport(ss, regularReportSheetName, regularReportNumDays, today, regexMap, regularShiftMap, latestDataMap, reservationData, specialistShiftMap);
    Logger.log(' -> 「予約リスト」の作成完了');

    // --- STEP 3: 専任シフトのレポートを作成 ---
    Logger.log('3. 「専任シフト予約リスト」を作成します...');
    const specialistReportData = createSpecialistReport(ss, specialistReportSheetName, specialistShiftMap, regexMap, latestDataMap, reservationData);
    Logger.log(' -> 「専任シフト予約リスト」の作成完了');

    // --- STEP 4: 未充足拠点のリストを作成 ---
    Logger.log('4. 「医師未充足拠点」リストを作成します...');
    createUnfulfilledReport(ss, unfulfilledSheetName, normalReportData, specialistReportData, today, regularReportNumDays);
    Logger.log(' -> 「医師未充足拠点」リストの作成完了');

    Logger.log('--- 全てのレポート更新が完了しました ---');

  } catch (e) {
    Logger.log(`🚨 重大なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * 通常の予約リストを作成・書き出しし、作成したデータを返す関数
 */
function createNormalReport(ss, sheetName, numDays, today, regexMap, shiftMap, latestDataMap, reservationData, specialistShiftMap) {
  const header = ['対象日', '判定', '拠点名', '診療科', 'ワクチン予約数', 'ワクチン予約枠数', '充足率', '午前医師', '午後医師', '夜間医師', '別会場'];
  const reportData = [];
  const processedKeys = new Set();

  for (let d = 0; d < numDays; d++) {
    const targetDate = new Date(today);
    targetDate.setDate(today.getDate() + d);
    const dateColumnIndex = reservationData[1].findIndex(dateCell => dateCell instanceof Date && dateCell.getTime() === targetDate.getTime());
    if (dateColumnIndex === -1) continue;

    for (let i = 3; i < reservationData.length; i++) {
      const locationName = reservationData[i][6];
      const department = reservationData[i][7];
      if (!locationName || locationName.toString().includes('全クリニック') || !department) continue;
      const regexInfo = regexMap[locationName];
      if (!regexInfo || (regexInfo.openingDate && regexInfo.openingDate > today)) continue;

      const uniqueKey = `${targetDate.getTime()}_${locationName}_${department}`;
      if (processedKeys.has(uniqueKey)) continue;
      processedKeys.add(uniqueKey);

      const reservationCount = reservationData[i][dateColumnIndex] || 0;
      const officialLocationName = regexInfo.officialName;
      const latestDataMapKey = `${officialLocationName}_${targetDate.getTime()}`;
      const slotCount = latestDataMap[latestDataMapKey] ? latestDataMap[latestDataMapKey].slotCount : 0;
      if (reservationCount === 0 && slotCount === 0) continue;

      let reportLocationName = locationName;
      if ((locationName === '北葛西' || locationName === '亀有')) {
        reportLocationName = `${locationName}（${department}）`;
      }
      let shiftMapKey;
      if (reportLocationName.includes('（')) {
        shiftMapKey = `${reportLocationName}_${targetDate.getTime()}`;
      } else {
        shiftMapKey = `${officialLocationName}_${targetDate.getTime()}`;
      }
      const shiftInfo = shiftMap[shiftMapKey] || { morning: 0, afternoon: 0, night: 0 };
      const fulfillmentRate = slotCount > 0 ? reservationCount / slotCount : 0;
      
      // ★★★ 修正点: 午前・午後・夜間の全てが1以上の場合のみ「充足」とする ★★★
      const judgment = (shiftInfo.morning > 0 && shiftInfo.afternoon > 0 && shiftInfo.night > 0) ? '充足' : '未充足';

      let specialistStatus = '';
      if (officialLocationName === '亀有' || officialLocationName === '柏の葉') {
        const specialistShiftKey = `${officialLocationName}_${targetDate.getTime()}`;
        const specialistShiftInfo = specialistShiftMap[specialistShiftKey];
        if (specialistShiftInfo) {
          if (specialistShiftInfo.morning > 0 && specialistShiftInfo.afternoon > 0) {
            specialistStatus = '充足';
          } else if (specialistShiftInfo.morning === 0 && specialistShiftInfo.afternoon > 0) {
            specialistStatus = '午前：不在';
          } else if (specialistShiftInfo.morning > 0 && specialistShiftInfo.afternoon === 0) {
            specialistStatus = '午後：不在';
          } else if (specialistShiftInfo.morning === 0 && specialistShiftInfo.afternoon === 0) {
            specialistStatus = '両方不在';
          }
        }
      }
      reportData.push([
        targetDate, judgment, reportLocationName, department,
        reservationCount, slotCount, fulfillmentRate,
        shiftInfo.morning, shiftInfo.afternoon, shiftInfo.night,
        specialistStatus
      ]);
    }
  }
  writeReportToSheet(ss, sheetName, header, reportData, false);
  return reportData;
}

/**
 * 専任シフトの予約リストを作成・書き出しし、作成したデータを返す関数
 */
function createSpecialistReport(ss, sheetName, specialistShiftMap, regexMap, latestDataMap, reservationData) {
  const reportData = [];
  const header = ['対象日', '判定', '拠点名', '診療科', 'ワクチン予約数', 'ワクチン予約枠数', '充足率', '午前医師', '午後医師', '夜間医師'];
  const cleanReservationMap = createCleanReservationMap(reservationData, regexMap);

  for (const shiftKey in specialistShiftMap) {
    const shiftInfo = specialistShiftMap[shiftKey];
    const [locationName, dateMilli] = shiftKey.split('_');
    const targetDate = new Date(parseInt(dateMilli));
    let reservationCount = 0;
    let department = '専任';
    for(const key in cleanReservationMap){
        const record = cleanReservationMap[key];
        if(record.officialLocationName === locationName && record.targetDate.getTime() === targetDate.getTime()){
            reservationCount += record.reservationCount;
            department = record.department;
        }
    }
    department = '専任';
    const officialLocationName = locationName;
    const latestDataMapKey = `${officialLocationName}_${targetDate.getTime()}`;
    const slotCount = latestDataMap[latestDataMapKey] ? latestDataMap[latestDataMapKey].slotCount : 0;
    if (reservationCount === 0 && slotCount === 0) continue;
    const judgment = (shiftInfo.morning > 0 && shiftInfo.afternoon > 0) ? '充足' : '未充足';
    const nightShiftDisplay = '未募集';
    const fulfillmentRate = slotCount > 0 ? reservationCount / slotCount : 0;
    reportData.push([
      targetDate, judgment, locationName, department,
      reservationCount, slotCount, fulfillmentRate,
      shiftInfo.morning, shiftInfo.afternoon, nightShiftDisplay
    ]);
  }
  writeReportToSheet(ss, sheetName, header, reportData, true);
  return reportData;
}

/**
 * 未充足拠点のレポートを作成・書き出しする関数
 */
function createUnfulfilledReport(ss, sheetName, normalReportData, specialistReportData, today, numDays) {
  const header = ['対象日', '判定', '拠点名', '診療科', 'ワクチン予約数', 'ワクチン予約枠数', '充足率', '午前医師', '午後医師', '夜間医師', '別会場'];
  const unfulfilledList = [];
  const endDate = new Date(today);
  endDate.setDate(today.getDate() + numDays);

  normalReportData.forEach(row => {
    if (row[1] === '未充足' && row[0] < endDate) {
      unfulfilledList.push(row);
    }
  });

  specialistReportData.forEach(row => {
    if (row[1] === '未充足' && row[0] < endDate) {
      const newRow = [...row];
      newRow.push('別会場');
      unfulfilledList.push(newRow);
    }
  });

  writeReportToSheet(ss, sheetName, header, unfulfilledList, false);
}

/**
 * レポートデータをシートに書き出す共通関数
 */
function writeReportToSheet(ss, sheetName, header, data, isSpecialistReport) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clear();
  }

  if (data.length > 0) {
    data.sort((a, b) => {
      if (a[0].getTime() !== b[0].getTime()) return a[0].getTime() - b[0].getTime();
      const nameA = a[2];
      const nameB = b[2];
      return nameA.localeCompare(nameB);
    });
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    const bodyRange = sheet.getRange(2, 1, data.length, header.length);
    bodyRange.setValues(data);
    bodyRange.setNumberFormat('@');
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setHorizontalAlignment('left');
    
    if (sheetName === '医師未充足拠点') {
        sheet.getRange(2, 1, data.length, 1).setNumberFormat('yyyy/mm/dd');
        sheet.getRange(2, 5, data.length, 2).setNumberFormat('0');
        sheet.getRange(2, 7, data.length, 1).setNumberFormat('0.0%');
        sheet.getRange(2, 8, data.length, 3).setNumberFormat('0');
    } else if (isSpecialistReport) {
        sheet.getRange(2, 1, data.length, 1).setNumberFormat('yyyy/mm/dd');
        sheet.getRange(2, 5, data.length, 2).setNumberFormat('0');
        sheet.getRange(2, 7, data.length, 1).setNumberFormat('0.0%');
        sheet.getRange(2, 8, data.length, 2).setNumberFormat('0');
    } else {
        sheet.getRange(2, 1, data.length, 1).setNumberFormat('yyyy/mm/dd');
        sheet.getRange(2, 5, data.length, 2).setNumberFormat('0');
        sheet.getRange(2, 7, data.length, 1).setNumberFormat('0.0%');
        sheet.getRange(2, 8, data.length, 3).setNumberFormat('0');
    }
  } else {
    Logger.log(`レポートデータがなかったため、「${sheetName}」は更新されませんでした。`);
  }
}