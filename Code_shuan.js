/**
 * 週案 Webアプリケーション化システム（完全版）
 */

function doGet(e) {
  const allowedPages = ['index', 'WeeklyPrint', 'tsushin', 'setting', 'ipad'];
  const requestedPage = e && e.parameter ? e.parameter.p : '';
  const page = allowedPages.includes(requestedPage) ? requestedPage : 'index';

  const appSettings = getAppDisplaySettings_();
  const template = HtmlService.createTemplateFromFile(page);
  template.appName = appSettings.appName;
  template.schoolName = appSettings.schoolName;
  template.className = appSettings.className;
  template.teacherName = appSettings.teacherName;

  return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle(appSettings.appName);
}

// 他の画面へのリンクを生成するための補助関数（HTML側で使用）
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function exportAsHtmlWeekly() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('WeeklyPrint')
      .setWidth(1000)
      .setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '週案管理システム');
}

function getCurrentWeekNum() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("週案") || ss.getSheets()[0];
  return sheet ? String(sheet.getRange("H2").getValue()) : "1";
}
// 今日の日付から「第何週目か」を自動計算して返す関数
function getActualCurrentWeekNum() {
  const weeks = getWeekConfigData_();
  const today = new Date();
  today.setHours(0,0,0,0); // 時間のズレをリセット
  
  let currentWeek = "1";
  
  // 設定シートの各週の期間と今日の日付を比較
  for (let i = 0; i < weeks.length; i++) {
    if (!weeks[i].startDate) continue;
    const startMonday = new Date(weeks[i].startDate);
    startMonday.setHours(0,0,0,0);
    
    // その週の日曜日を計算
    const endSunday = new Date(startMonday);
    endSunday.setDate(startMonday.getDate() + 6);
    endSunday.setHours(23,59,59,999);
    
    // 今日が「月曜〜日曜」の間に挟まっていれば、その週番号を返す
    if (today >= startMonday && today <= endSunday) {
      currentWeek = String(weeks[i].weekNum);
      break;
    }
  }
  return currentWeek;
}

function getWeekConfigData_() {
  const cache = CacheService.getDocumentCache();
  const key = 'WEEK_CONFIG_V1';
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  if (!sheet) return [];
  const result = sheet.getRange('O9:P60').getValues().map(function(row) {
    return { weekNum: Number(row[0]), startDate: row[1] instanceof Date && !isNaN(row[1].getTime()) ? row[1].toISOString() : '' };
  }).filter(function(item) { return item.weekNum && item.startDate; });
  cache.put(key, JSON.stringify(result), 600);
  return result;
}
// 設定シートから基本時間割（U2:Z8）を取得する関数
function getDefaultTimetable() {
  const cached = getWeeklyStaticSettings_();
  if (cached && cached.defaultTimetable) return cached.defaultTimetable;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  if (!configSheet) return { error: "設定シートが見つかりません。" };
  
  const defaultData = configSheet.getRange("U2:Z8").getValues();
  const defaultTimetable = { p1: [], p2: [], p3: [], p4: [], p5: [], p6: [] };
  const dayIndices = [1, 2, 3, 4, 5]; // V列(月)=1, W列(火)=2, X列(水)=3, Y列(木)=4, Z列(金)=5
  
  for (let d = 0; d < 7; d++) {
    if (d < 5) {
      const col = dayIndices[d];
      defaultTimetable.p1.push(String(defaultData[1][col] || ""));
      defaultTimetable.p2.push(String(defaultData[2][col] || ""));
      defaultTimetable.p3.push(String(defaultData[3][col] || ""));
      defaultTimetable.p4.push(String(defaultData[4][col] || ""));
      defaultTimetable.p5.push(String(defaultData[5][col] || ""));
      defaultTimetable.p6.push(String(defaultData[6][col] || ""));
    } else {
      defaultTimetable.p1.push(""); defaultTimetable.p2.push(""); defaultTimetable.p3.push("");
      defaultTimetable.p4.push(""); defaultTimetable.p5.push(""); defaultTimetable.p6.push("");
    }
  }
  return defaultTimetable;
}

/** 変更頻度の低い週案設定を短時間キャッシュする。 */
function getWeeklyStaticSettings_() {
  const cache = CacheService.getDocumentCache();
  const key = 'WEEKLY_STATIC_V1';
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  if (!sheet) return null;
  const values = sheet.getRange('L3:AB22').getValues();
  const subjects = values.slice(5, 20).map(r => String(r[0] || '').trim())
    .filter(v => v && v !== '教科');
  const leavingTimes = values.slice(5, 20).map(r => {
    const value = r[1];
    return value instanceof Date
      ? value.getHours() + ':' + String(value.getMinutes()).padStart(2, '0')
      : String(value || '').trim();
  }).filter(v => v && v !== '下校時間');
  const colorSubjects = values.slice(0, 10).map(r => String(r[16] || '').trim())
    .filter(v => v && !v.includes('色を変える教科'));
  const source = sheet.getRange('U2:Z8').getValues();
  const defaultTimetable = { p1: [], p2: [], p3: [], p4: [], p5: [], p6: [] };
  for (let p = 1; p <= 6; p++) {
    for (let d = 0; d < 7; d++) {
      defaultTimetable['p' + p].push(d < 5 ? String(source[p][d + 1] || '') : '');
    }
  }
  const result = { subjects, leavingTimes, colorSubjects, defaultTimetable };
  cache.put(key, JSON.stringify(result), 600);
  return result;
}

function clearShuanCaches_() {
  const cache = CacheService.getDocumentCache();
  cache.remove('WEEKLY_STATIC_V1');
  cache.remove('WEEKLY_DATE_MAP_V1');
  cache.remove('WEEK_CONFIG_V1');
  cache.remove('SETTINGS_DATA_V1');
}

/** 年間の日付行だけを読み、日付から列番号を引けるようキャッシュする。 */
function getWeeklyDateColumnMap_(sheet) {
  const cache = CacheService.getDocumentCache();
  const key = 'WEEKLY_DATE_MAP_V1';
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }
  const lastCol = sheet.getLastColumn();
  const dates = lastCol >= 2 ? sheet.getRange(10, 2, 1, lastCol - 1).getValues()[0] : [];
  const map = {};
  dates.forEach((value, i) => {
    if (value instanceof Date && !isNaN(value.getTime())) {
      map[Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')] = i + 2;
    }
  });
  cache.put(key, JSON.stringify(map), 600);
  return map;
}

// データを読み込むメイン関数
function getWeeklyDataByNumber(targetWeekNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("週案") || ss.getSheets()[0]; 
  const configSheet = ss.getSheetByName("設定");
  
  if (!mainSheet || !configSheet) return { error: "シート名を確認してください。" };

  const targetWeek = getWeekConfigData_().find(item => Number(item.weekNum) === Number(targetWeekNum));
  if (!targetWeek) return { error: "指定された週が見つかりません。" };
  const startMonday = new Date(targetWeek.startDate);
  const weekDates = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(startMonday);
    d.setDate(startMonday.getDate() + i);
    weekDates.push(d);
  }
  const dayLabels = ["月", "火", "水", "木", "金", "土", "日"];

  const dateMap = getWeeklyDateColumnMap_(mainSheet);
  const sheetColumns = weekDates.map(date => dateMap[Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd')] || -1);
  if (sheetColumns.some(col => col < 1)) return { error: '対象週の日付列が見つかりません。' };
  const firstColumn = sheetColumns[0];
  if (!sheetColumns.every((col, i) => col === firstColumn + i)) return { error: '対象週の日付列が連続していません。' };
  const mainData = mainSheet.getRange(1, firstColumn, 35, 7).getValues();
  const colIndices = sheetColumns.map(col => col - 1); // クライアント互換の0始まり
  const staticSettings = getWeeklyStaticSettings_() || { subjects: [], colorSubjects: [], leavingTimes: [] };

  const getSplitRowCells = (rowNum1, rowNum2) => {
    return sheetColumns.map((col, localCol) => {
      return {
        subject: String(mainData[rowNum1-1][localCol] || "").trim(),
        content: String(mainData[rowNum2-1][localCol] || "").trim()
      };
    });
  };

  const getRowCells = (rowNum) => {
    return sheetColumns.map((col, localCol) => {
      return String(mainData[rowNum-1][localCol] || "");
    });
  };

  const getLeavingRowCells = (rowNum) => {
    return sheetColumns.map((col, localCol) => {
      const val = mainData[rowNum-1][localCol];
      if (val instanceof Date) {
        return val.getHours() + ":" + String(val.getMinutes()).padStart(2, '0');
      }
      return String(val || "");
    });
  };

  const appSettings = getAppDisplaySettings_();
  return {
    year: String(mainSheet.getRange('A2').getValue()).replace(/年度/g, ""),
    schoolName: appSettings.schoolName,
    className: appSettings.className,
    teacherName: appSettings.teacherName,
    startDate: (weekDates[0].getMonth() + 1) + "月" + weekDates[0].getDate() + "日",
    endDate: (weekDates[6].getMonth() + 1) + "月" + weekDates[6].getDate() + "日",
    weekNum: targetWeekNum,
    days: weekDates.map((d, idx) => `${d.getMonth() + 1}/${d.getDate()}(${dayLabels[idx]})`),
    holidayRow: getRowCells(7),
    dropdownSubjects: staticSettings.subjects,
    colorSubjects: staticSettings.colorSubjects,
    leavingTimes: staticSettings.leavingTimes,
    colIndices: colIndices,
    
    timetable: {
      event:    getRowCells(11),
      morning:  getRowCells(12),
      p1:       getSplitRowCells(13, 14),
      p2:       getSplitRowCells(15, 16),
      rest:     getRowCells(17),
      p3:       getSplitRowCells(18, 19),
      p4:       getSplitRowCells(20, 21),
      lunch:    getRowCells(22),
      clean:    getRowCells(23),
      recess:   getRowCells(24),
      p5:       getSplitRowCells(25, 26),
      p6:       getSplitRowCells(27, 28),
      leaving:  getLeavingRowCells(29),
      homework: getRowCells(30),
      memo:     getRowCells(31)
    }
  };
}

function saveWeeklyData(weekNum, colIndices, clientData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("週案") || ss.getSheets()[0];
  if (!mainSheet) return { error: "シートが見つかりません。" };

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    const columns = (colIndices || []).map(Number);
    if (columns.length !== 7 || columns.some((col, i) => !Number.isInteger(col) || col < 0 || (i && col !== columns[0] + i))) {
      return { error: '保存対象の列情報が正しくありません。' };
    }
    const rowKeys = [
      ['single','morning'], ['subject','p1'], ['content','p1'], ['subject','p2'], ['content','p2'],
      ['single','rest'], ['subject','p3'], ['content','p3'], ['subject','p4'], ['content','p4'],
      ['single','lunch'], ['single','clean'], ['single','recess'], ['subject','p5'], ['content','p5'],
      ['subject','p6'], ['content','p6'], ['single','leaving'], ['single','homework'], ['single','memo']
    ];
    const values = rowKeys.map(([kind, key]) => Array.from({ length: 7 }, (_, i) => {
      if (kind === 'single') return (clientData[key] || [])[i] || '';
      const cell = (clientData[key] || [])[i] || {};
      return cell[kind] || '';
    }));
    mainSheet.getRange(12, columns[0] + 1, 20, 7).setValues(values);
    PropertiesService.getUserProperties().deleteProperty('WEEKLY_DRAFT_' + String(weekNum));
    return { success: true };
  } catch(e) {
    return { error: e.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function getWeeklyPageInitialData() {
  const weekNum = Number(getActualCurrentWeekNum());
  return {
    weekNum: weekNum,
    weekly: getWeeklyDataByNumber(weekNum),
    slideUrl: getSlideUrl(),
    productivity: getProductivitySettings(),
    defaultTimetable: getDefaultTimetable(),
    draft: getWeeklyDraft(weekNum)
  };
}

function getWeeklyPageData(weekNum) {
  return { weekly: getWeeklyDataByNumber(weekNum), draft: getWeeklyDraft(weekNum) };
}

// Webアプリから行事・休日を設定シートの最後に追加する関数
function addNewConfigEventOrHoliday(type, dateStr, nameStr) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("設定");
    if (!configSheet) return { error: "設定シートが見つかりません。" };

    if (type !== 'event' && type !== 'holiday') {
      return { error: "登録種別が正しくありません。" };
    }

    const name = String(nameStr || "").trim();
    if (!name) return { error: "名称を入力してください。" };

    const parts = String(dateStr || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!parts) return { error: "日付の形式が正しくありません。" };

    const targetDate = new Date(Number(parts[1]), Number(parts[2]) - 1, Number(parts[3]));
    if (targetDate.getFullYear() !== Number(parts[1]) ||
        targetDate.getMonth() !== Number(parts[2]) - 1 ||
        targetDate.getDate() !== Number(parts[3])) {
      return { error: "存在しない日付です。" };
    }

    const dateCol = type === 'event' ? 6 : 8;
    const nameCol = dateCol + 1;
    const startRow = 3;
    const lastRow = Math.max(configSheet.getLastRow(), startRow - 1);
    const rowCount = Math.max(lastRow - startRow + 1, 0);
    const existing = rowCount > 0
      ? configSheet.getRange(startRow, dateCol, rowCount, 2).getValues()
      : [];

    const targetTime = targetDate.getTime();
    const duplicate = existing.some(row => {
      if (!(row[0] instanceof Date)) return false;
      const d = new Date(row[0]);
      d.setHours(0, 0, 0, 0);
      return d.getTime() === targetTime && String(row[1] || "").trim() === name;
    });
    if (duplicate) {
      return { error: "同じ日付・名称がすでに登録されています。" };
    }

    let nextRow = startRow;
    for (let i = 0; i < existing.length; i++) {
      if (existing[i][0] !== "" || existing[i][1] !== "") nextRow = startRow + i + 1;
    }

    configSheet.getRange(nextRow, dateCol, 1, 2)
      .setValues([[targetDate, name]]);
    configSheet.getRange(nextRow, dateCol).setNumberFormat("yyyy/MM/dd");
    SpreadsheetApp.flush();

    const label = type === 'event' ? "行事" : "休日";
    return { success: true, message: label + "を登録しました。" };
  } catch (e) {
    return { error: "登録に失敗しました: " + e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function getSlideUrl() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("設定");
    if (!sheet) {
      // 万が一「設定」シートという名前がない場合は、左端のシートを参照
      sheet = ss.getSheets()[0];
    }
    var url = sheet.getRange("U13").getValue();
    return url;
  } catch (e) {
    Logger.log("エラー: " + e.toString());
    return "";
  }
}
