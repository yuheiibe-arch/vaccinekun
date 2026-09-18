/**
 * =======================================================================================
 * 共通の補助機能ファイル
 * =======================================================================================
 * ・キャッシュ処理
 * ・「データ」シートからの情報取得
 */

const SCRIPT_VERSION = 'v21'; // バージョン更新
const CACHE_EXPIRATION = 21600; // 6時間

/**
 * キャッシュからデータを取得、なければ関数を実行して取得・保存する汎用関数
 */
function getCachedData(key, fetchFunction) {
  const versionedKey = `${SCRIPT_VERSION}_${key}`;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(versionedKey);
  if (cached != null) {
    // Logger.log(`[キャッシュ] ${versionedKey} のデータをキャッシュから取得しました。`);
    return JSON.parse(cached, (k, v) => (typeof v === 'string' && v.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) ? new Date(v) : v);
  }
  Logger.log(`[シートから読込] ${versionedKey} のデータをシートから読み込み、キャッシュに保存します...`);
  const data = fetchFunction();
  cache.put(versionedKey, JSON.stringify(data), CACHE_EXPIRATION);
  return data;
}

/**
 * 「データ」シートから、対象となるシートのURLとシート名を取得する
 */
function getSheetInfo(masterSheetName, targetName) {
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterSheetName);
  if (!masterSheet) throw new Error(`「${masterSheetName}」シートが見つかりません。`);
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][0] === targetName) {
      return { url: masterData[i][1], sheetName: masterData[i][2] || targetName };
    }
  }
  throw new Error(`「${masterSheetName}」シートに「${targetName}」の情報が見つかりませんでした。`);
}