/**
 * =======================================================================================
 * データ取得・加工ファイル（時間ベース完全充足判定・キャッシュ容量オーバー回避版）
 * =======================================================================================
 */

// --- キャッシュラッパー関数（修正部分） ---
function getRegexMap(masterSheetName, targetName) {
  return getCachedData('regexMap_v6', () => createRegexMap(masterSheetName, targetName));
}

// ★ 修正：データ量がGoogleの上限（100KB）を超えるため、一時保存を使わず毎回直接読み込みます
function getShiftMap(masterSheetName, targetName, numDays, regexMap) {
  return createShiftMapForPeriod(masterSheetName, targetName, numDays, regexMap);
}

// ★ 同様にデータ量オーバーによるクラッシュを防ぐため一時保存を外します
function getLatestDataMap(masterSheetName, urlLookupName, targetHistorySheetName, numDays, regexMap) {
  return createLatestDataMap(masterSheetName, urlLookupName, targetHistorySheetName, numDays, regexMap);
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

  const timeToMinutes = (timeInput) => {
    if (timeInput === null || timeInput === undefined || timeInput === '') return NaN;
    if (timeInput instanceof Date) return timeInput.getHours() * 60 + timeInput.getMinutes();
    if (typeof timeInput === 'string') {
      const parts = timeInput.trim().split(':');
      if (parts.length >= 2) return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
    }
    if (typeof timeInput === 'number' && timeInput >= 0 && timeInput <= 1) return Math.round(timeInput * 24 * 60);
    return NaN;
  };

  const workDataTemp = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const doctorName = row[0] ? String(row[0]).trim() : "";
    const clinicName = row[12];
    const department = row[13];
    const dateCell = row[14];
    
    if (doctorName.includes("バックアップ") || doctorName.includes("有給") || doctorName.includes("欠勤")) continue;
    if (!clinicName || !dateCell) continue;
    
    let shiftDateObj;
    if (dateCell instanceof Date) {
      shiftDateObj = dateCell;
    } else if (typeof dateCell === 'string' && dateCell.trim() !== '') {
      shiftDateObj = new Date(dateCell.split('（')[0].trim().replace(/-/g, '/'));
    } else {
      continue;
    }

    if (isNaN(shiftDateObj.getTime()) || !(shiftDateObj >= today && shiftDateObj < endDate)) continue;
    
    const regexInfo = regexMap[clinicName];
    if (!regexInfo) continue;
    
    let keyName = regexInfo.officialName;
    if (keyName === '北葛西' || keyName === '亀有') {
      keyName = `${keyName}（${department}）`;
    }
    
    const mapKey = `${keyName}_${shiftDateObj.getTime()}`;
    if (!workDataTemp[mapKey]) {
      workDataTemp[mapKey] = { intervals: [] };
    }
    
    const startMin = timeToMinutes(row[15]);
    const endMin = timeToMinutes(row[19]);
    if (!isNaN(startMin) && !isNaN(endMin) && startMin < endMin) {
      workDataTemp[mapKey].intervals.push({ start: startMin, end: endMin });
    }
  }

  const mergeIntervals = (intervals) => {
    if (!intervals || intervals.length === 0) return [];
    intervals.sort((a, b) => a.start - b.start);
    const merged = [];
    let current = { ...intervals[0] };
    for (let i = 1; i < intervals.length; i++) {
      if (intervals[i].start <= current.end) {
        current.end = Math.max(current.end, intervals[i].end);
      } else {
        merged.push(current);
        current = { ...intervals[i] };
      }
    }
    merged.push(current);
    return merged;
  };
  
  const formatMinutes = (m) => `${String(Math.floor(m/60)).padStart(2,'0')}:${String(m%60).padStart(2,'0')}`;
  const isOverlap = (gapStart, gapEnd, checkStart, checkEnd) => gapStart < checkEnd && gapEnd > checkStart;

  for (const mapKey in workDataTemp) {
    const temp = workDataTemp[mapKey];
    const merged = mergeIntervals(temp.intervals);
    
    let clinicClose = 21 * 60;
    if (mapKey.includes('北葛西')) clinicClose = 20 * 60;

    const requiredShifts = [
      { start: 9 * 60, end: 13 * 60 },
      { start: 15 * 60, end: clinicClose }
    ];
    
    let absenceGaps = [];
    let isMorningAbsent = false;
    let isAfternoonAbsent = false;
    let isNightAbsent = false;
    
    if (merged.length === 0) {
        isMorningAbsent = true;
        isAfternoonAbsent = true;
        isNightAbsent = true;
    } else {
        requiredShifts.forEach(req => {
          let currentTimePointer = req.start;
          for (const work of merged) {
            if (currentTimePointer < work.start && currentTimePointer < req.end) {
              const gapEnd = Math.min(work.start, req.end);
              absenceGaps.push(`${formatMinutes(currentTimePointer)}-${formatMinutes(gapEnd)}`);
              
              if (isOverlap(currentTimePointer, gapEnd, 9 * 60, 13 * 60)) isMorningAbsent = true;
              if (isOverlap(currentTimePointer, gapEnd, 15 * 60, 18 * 60)) isAfternoonAbsent = true;
              if (isOverlap(currentTimePointer, gapEnd, 18 * 60, clinicClose)) isNightAbsent = true;
            }
            currentTimePointer = Math.max(currentTimePointer, work.end);
          }
          if (currentTimePointer < req.end) {
            const gapEnd = req.end;
            absenceGaps.push(`${formatMinutes(currentTimePointer)}-${formatMinutes(gapEnd)}`);
            
            if (isOverlap(currentTimePointer, gapEnd, 9 * 60, 13 * 60)) isMorningAbsent = true;
            if (isOverlap(currentTimePointer, gapEnd, 15 * 60, 18 * 60)) isAfternoonAbsent = true;
            if (isOverlap(currentTimePointer, gapEnd, 18 * 60, clinicClose)) isNightAbsent = true;
          }
        });
    }

    map[mapKey] = {
      morning: isMorningAbsent ? 0 : 1,
      afternoon: isAfternoonAbsent ? 0 : 1,
      night: isNightAbsent ? 0 : 1,
      absenceDetails: absenceGaps.length > 0 ? `（${absenceGaps.join(', ')}で医師不在あり）` : ""
    };
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
  
  const TARGET_CLINICS = ["亀有", "柏の葉"];
  const TARGET_DEPARTMENTS = ["小児科ワクチン専任(対象：小児～成人)", "内科ワクチン専任(対象：小児～成人)"];
  
  const timeToMinutes = (timeInput) => {
    if (timeInput === null || timeInput === undefined || timeInput === '') return NaN;
    if (timeInput instanceof Date) return timeInput.getHours() * 60 + timeInput.getMinutes();
    if (typeof timeInput === 'string') {
      const parts = timeInput.trim().split(':');
      if (parts.length >= 2) return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
    }
    if (typeof timeInput === 'number' && timeInput >= 0 && timeInput <= 1) return Math.round(timeInput * 24 * 60);
    return NaN;
  };

  const workDataTemp = {};
  const map = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const doctorName = row[0] ? String(row[0]).trim() : "";
    const originalClinicName = row[12] ? row[12].toString().trim() : null;
    const originalDepartment = row[13] ? row[13].toString().trim() : null;
    
    if (doctorName.includes("バックアップ") || doctorName.includes("有給") || doctorName.includes("欠勤")) continue;
    if (!TARGET_DEPARTMENTS.includes(originalDepartment) || !TARGET_CLINICS.includes(originalClinicName)) continue;
    
    const shiftDateRaw = row[14];
    if (!shiftDateRaw) continue;
    let shiftDateObj;
    
    try {
      if (shiftDateRaw instanceof Date) {
        shiftDateObj = shiftDateRaw;
      } else {
        shiftDateObj = new Date(shiftDateRaw.toString().split('（')[0].trim().replace(/-/g, '/'));
      }
      if (isNaN(shiftDateObj.getTime())) continue;
      shiftDateObj.setHours(0, 0, 0, 0);
    } catch (e) { continue; }
    
    const regexInfo = regexMap[originalClinicName];
    if (!regexInfo) continue;
    const officialLocationName = regexInfo.officialName;
    const mapKey = `${officialLocationName}_${shiftDateObj.getTime()}`;
    
    if (!workDataTemp[mapKey]) workDataTemp[mapKey] = { intervals: [] };
    
    const startMin = timeToMinutes(row[15]);
    const endMin = timeToMinutes(row[19]);
    if (!isNaN(startMin) && !isNaN(endMin) && startMin < endMin) {
      workDataTemp[mapKey].intervals.push({ start: startMin, end: endMin });
    }
  }

  const mergeIntervals = (intervals) => {
    if (!intervals || intervals.length === 0) return [];
    intervals.sort((a, b) => a.start - b.start);
    const merged = [];
    let current = { ...intervals[0] };
    for (let i = 1; i < intervals.length; i++) {
      if (intervals[i].start <= current.end) {
        current.end = Math.max(current.end, intervals[i].end);
      } else {
        merged.push(current);
        current = { ...intervals[i] };
      }
    }
    merged.push(current);
    return merged;
  };

  const isOverlap = (gapStart, gapEnd, checkStart, checkEnd) => gapStart < checkEnd && gapEnd > checkStart;

  for (const mapKey in workDataTemp) {
    const merged = mergeIntervals(workDataTemp[mapKey].intervals);
    
    let isMorningAbsent = false;
    let isAfternoonAbsent = false;
    
    const requiredShifts = [
      { start: 9 * 60, end: 13 * 60 },
      { start: 15 * 60, end: 18 * 60 }
    ];
    
    if (merged.length === 0) {
        isMorningAbsent = true;
        isAfternoonAbsent = true;
    } else {
        requiredShifts.forEach(req => {
          let currentTimePointer = req.start;
          for (const work of merged) {
            if (currentTimePointer < work.start && currentTimePointer < req.end) {
              const gapEnd = Math.min(work.start, req.end);
              if (isOverlap(currentTimePointer, gapEnd, 9 * 60, 13 * 60)) isMorningAbsent = true;
              if (isOverlap(currentTimePointer, gapEnd, 15 * 60, 18 * 60)) isAfternoonAbsent = true;
            }
            currentTimePointer = Math.max(currentTimePointer, work.end);
          }
          if (currentTimePointer < req.end) {
            const gapEnd = req.end;
            if (isOverlap(currentTimePointer, gapEnd, 9 * 60, 13 * 60)) isMorningAbsent = true;
            if (isOverlap(currentTimePointer, gapEnd, 15 * 60, 18 * 60)) isAfternoonAbsent = true;
          }
        });
    }

    map[mapKey] = {
      morning: isMorningAbsent ? 0 : 1,
      afternoon: isAfternoonAbsent ? 0 : 1,
      night: 1
    };
  }
  
  return map;
}