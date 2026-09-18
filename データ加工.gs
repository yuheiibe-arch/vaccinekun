/**
 * =======================================================================================
 * データ取得・加工ファイル（完成版・日付文字列の解析に対応）
 * =======================================================================================
 * ・各種マスタデータ、シフトデータ等をスプレッドシートから読み込み、
 * 　レポートで使いやすい「マップ」形式に加工する。
 */

// --- キャッシュラッパー関数 ---

function getRegexMap(masterSheetName, targetName) {
  return getCachedData('regexMap', () => createRegexMap(masterSheetName, targetName));
}

function getShiftMap(masterSheetName, targetName, numDays, regexMap) {
  return getCachedData('shiftMap', () => createShiftMapForPeriod(masterSheetName, targetName, numDays, regexMap));
}

function getLatestDataMap(masterSheetName, urlLookupName, targetHistorySheetName, numDays, regexMap) {
  return getCachedData('latestDataMap', () => createLatestDataMap(masterSheetName, urlLookupName, targetHistorySheetName, numDays, regexMap));
}


// --- データソース読み込み・整形関数 ---

function createCleanReservationMap(reservationData, regexMap) {
  const cleanMap = {};
  const headerRow = reservationData[1];

  for (let i = 3; i < reservationData.length; i++) {
    const locationName = reservationData[i][6];
    const department = reservationData[i][7];
    if (!locationName || !department) continue;
    
    const regexInfo = regexMap[locationName];
    if (!regexInfo) continue;
    
    for (let c = 0; c < headerRow.length; c++) {
      const targetDate = headerRow[c];
      if (!(targetDate instanceof Date)) continue;

      const uniqueKey = `${targetDate.getTime()}_${locationName}_${department}`;
      if (!cleanMap[uniqueKey]) {
        cleanMap[uniqueKey] = {
          targetDate: targetDate,
          locationName: locationName,
          department: department,
          officialLocationName: regexInfo.officialName,
          reservationCount: reservationData[i][c] || 0,
        };
      }
    }
  }
  return cleanMap;
}


function createRegexMap(masterSheetName, targetName) {
  const info = getSheetInfo(masterSheetName, targetName);
  const sheet = SpreadsheetApp.openByUrl(info.url).getSheetByName(info.sheetName);
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const officialName = row[0];
    if (!officialName) continue;
    const alternativeNames = [row[1], row[2], row[3], row[4]];
    if (officialName === "北葛西") alternativeNames.push("北葛西（内科）", "北葛西（小児科）");
    if (officialName === "亀有") alternativeNames.push("亀有（内科）", "亀有（小児科）");
    if (officialName === "流山おおたかの森") alternativeNames.push("流山");
    const area = row[6];
    const openingDate = row[7] ? new Date(row[7]) : null;
    if (openingDate && !isNaN(openingDate.getTime())) openingDate.setHours(0, 0, 0, 0);
    const entry = { officialName: officialName, area: area, openingDate: openingDate };
    map[officialName] = entry;
    [...new Set(alternativeNames)].forEach(name => { if (name) map[name] = entry; });
  }
  return map;
}

function createShiftMapForPeriod(masterSheetName, targetName, numDays, regexMap) {
  const info = getSheetInfo(masterSheetName, targetName);
  const sheet = SpreadsheetApp.openByUrl(info.url).getSheetByName(info.sheetName);
  const data = sheet.getDataRange().getValues();
  const map = {};
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const endDate = new Date(today);
  endDate.setDate(today.getDate() + numDays);
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // ★★ 修正点: 日付の読み取り処理を、文字列と日付オブジェクトの両方に対応するロジックに変更 ★★
    const dateCell = row[1];
    let shiftDate;

    if (dateCell instanceof Date) {
      // セルが日付オブジェクトの場合
      shiftDate = dateCell;
    } else if (typeof dateCell === 'string' && dateCell.trim() !== '') {
      // セルが文字列の場合、解析を試みる
      const dateStringOnly = dateCell.split('（')[0].trim();
      shiftDate = new Date(dateStringOnly);
    } else {
      // 空白や無効なデータ型の場合は、この行をスキップ
      continue;
    }
    // ★★ 修正ここまで ★★

    if (isNaN(shiftDate.getTime()) || !(shiftDate >= today && shiftDate < endDate)) continue;
    const rawLocationName = row[0];
    const regexInfo = regexMap[rawLocationName];
    if (!regexInfo) continue;
    let keyName = (rawLocationName.includes('北葛西（') || rawLocationName.includes('亀有（')) ? rawLocationName : regexInfo.officialName;
    const mapKey = `${keyName}_${shiftDate.getTime()}`;
    const shiftData = map[mapKey] || { morning: 0, afternoon: 0, night: 0 };
    shiftData.morning += Number(row[3]) || 0;
    shiftData.afternoon += Number(row[4]) || 0;
    shiftData.night += Number(row[5]) || 0;
    map[mapKey] = shiftData;
  }
  return map;
}

function createLatestDataMap(masterSheetName, urlLookupName, targetHistorySheetName, numDays, regexMap) {
  const info = getSheetInfo(masterSheetName, urlLookupName);
  const spreadsheet = SpreadsheetApp.openByUrl(info.url);
  const sheet = spreadsheet.getSheetByName(targetHistorySheetName);
  if (!sheet) throw new Error(`シート「${targetHistorySheetName}」が見つかりません。`);
  const data = sheet.getDataRange().getValues();
  const map = {};
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const endDate = new Date(today);
  endDate.setDate(today.getDate() + numDays);
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[4]) continue;
    const reservationDate = new Date(row[4]);
    if (isNaN(reservationDate.getTime()) || !(reservationDate >= today && reservationDate < endDate)) continue;
    const rawLocationName = row[1];
    if (!rawLocationName) continue;
    const regexInfo = regexMap[rawLocationName];
    if (!regexInfo) continue;
    const officialLocationName = regexInfo.officialName;
    const mapKey = `${officialLocationName}_${reservationDate.getTime()}`;
    map[mapKey] = { slotCount: row[5] };
  }
  return map;
}

function createSpecialistShiftMap(spreadsheetId, sheetName, regexMap) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(2);
  const TARGET_CLINICS = ["亀有", "柏の葉"];
  const TARGET_DEPARTMENTS = ["小児科ワクチン専任(対象：小児～成人)", "内科ワクチン専任(対象：小児～成人)"];
  const map = {};
  rows.forEach((row) => {
    const originalDepartment = row[13] ? row[13].toString().trim() : null;
    const originalClinicName = row[12] ? row[12].toString().trim() : null;
    if (!TARGET_DEPARTMENTS.includes(originalDepartment) || !TARGET_CLINICS.includes(originalClinicName)) return;
    const shiftDateRaw = row[14];
    if (!shiftDateRaw) return;
    let shiftDateObj;
    try {
      shiftDateObj = new Date(shiftDateRaw.toString().split('（')[0].trim().replace(/-/g, '/'));
      if (isNaN(shiftDateObj.getTime())) return;
      shiftDateObj.setHours(0, 0, 0, 0);
    } catch (e) { return; }
    const regexInfo = regexMap[originalClinicName];
    if (!regexInfo) return;
    const officialLocationName = regexInfo.officialName;
    const mapKey = `${officialLocationName}_${shiftDateObj.getTime()}`;
    if (!map[mapKey]) map[mapKey] = { morning: 0, afternoon: 0, night: 0 };
    map[mapKey].morning += Number(row[55]) || 0;
    map[mapKey].afternoon += Number(row[56]) || 0;
    map[mapKey].night += Number(row[57]) || 0;
  });
  return map;
}