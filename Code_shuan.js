/**
 * 週案 Webアプリケーション化システム（完全版）
 */

function doGet(e) {
  // URLの末尾に「?p=ファイル名」があればそれを開き、なければ「index」を開く
  var page = e.parameter.p || 'index';
  
  return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('学級経営管理ツール');
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  if (!configSheet) return "1";

  const oValues = configSheet.getRange("O9:O60").getValues().map(r => r[0]); // 週番号(1〜52)
  const pValues = configSheet.getRange("P9:P60").getValues().map(r => r[0]); // 各週の月曜日の日付
  
  const today = new Date();
  today.setHours(0,0,0,0); // 時間のズレをリセット
  
  let currentWeek = "1";
  
  // 設定シートの各週の期間と今日の日付を比較
  for (let i = 0; i < pValues.length; i++) {
    if (!pValues[i]) continue;
    
    const startMonday = new Date(pValues[i]);
    startMonday.setHours(0,0,0,0);
    
    // その週の日曜日を計算
    const endSunday = new Date(startMonday);
    endSunday.setDate(startMonday.getDate() + 6);
    endSunday.setHours(23,59,59,999);
    
    // 今日が「月曜〜日曜」の間に挟まっていれば、その週番号を返す
    if (today >= startMonday && today <= endSunday) {
      currentWeek = String(oValues[i]);
      break;
    }
  }
  return currentWeek;
}
// 設定シートから基本時間割（U2:Z8）を取得する関数
function getDefaultTimetable() {
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

// データを読み込むメイン関数
function getWeeklyDataByNumber(targetWeekNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("週案") || ss.getSheets()[0]; 
  const configSheet = ss.getSheetByName("設定");
  
  if (!mainSheet || !configSheet) return { error: "シート名を確認してください。" };

  const oValues = configSheet.getRange("O9:O60").getValues().map(r => r[0]);
  const pValues = configSheet.getRange("P9:P60").getValues().map(r => r[0]);
  let targetIndex = oValues.indexOf(Number(targetWeekNum));
  if (targetIndex === -1) targetIndex = oValues.indexOf(String(targetWeekNum));
  if (targetIndex === -1) return { error: "指定された週が見つかりません。" };
  
  const startMonday = new Date(pValues[targetIndex]);
  const weekDates = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(startMonday);
    d.setDate(startMonday.getDate() + i);
    weekDates.push(d);
  }
  const dayLabels = ["月", "火", "水", "木", "金", "土", "日"];

  const lastCol = mainSheet.getLastColumn();
  const mainData = mainSheet.getRange(1, 1, 35, lastCol).getValues();
  const colIndices = [];
  weekDates.forEach((targetDate) => {
    let foundIndex = -1;
    for (let col = 1; col < lastCol; col++) {
      const cellVal = mainData[9][col];
      if (cellVal instanceof Date && cellVal.getMonth() === targetDate.getMonth() && cellVal.getDate() === targetDate.getDate()) {
        foundIndex = col; break;
      }
    }
    colIndices.push(foundIndex);
  });

  const lastRow = configSheet.getLastRow();

  // L列の8行目以降から「プルダウン用」の教科リストを取得
  let allSubjects = [];
  if (lastRow >= 8) {
    allSubjects = configSheet.getRange(8, 12, lastRow - 7, 1)
      .getValues()
      .map(r => String(r[0]).trim())
      .filter(val => val !== "" && val !== "undefined" && val !== "教科");
  }

  // AB列の3行目以降から「色を変える対象」の教科リストを取得
  let colorSubjects = [];
  if (lastRow >= 3) {
    colorSubjects = configSheet.getRange(3, 28, lastRow - 2, 1)
      .getValues()
      .map(r => String(r[0]).trim())
      .filter(val => val !== "" && val !== "undefined" && !val.includes("色を変える教科"));
  }

  // M列の8行目以降から下校時間リストを取得
  let leavingTimes = [];
  if (lastRow >= 8) {
    leavingTimes = configSheet.getRange(8, 13, lastRow - 7, 1)
      .getValues()
      .map(r => {
        const val = r[0];
        if (val instanceof Date) {
          return val.getHours() + ":" + String(val.getMinutes()).padStart(2, '0');
        }
        return String(val).trim();
      })
      .filter(val => val !== "" && val !== "undefined" && val !== "下校時間");
  }

  const getSplitRowCells = (rowNum1, rowNum2) => {
    return colIndices.map(col => {
      if (col === -1) return { subject: "", content: "" };
      return {
        subject: String(mainData[rowNum1-1][col] || "").trim(),
        content: String(mainData[rowNum2-1][col] || "").trim()
      };
    });
  };

  const getRowCells = (rowNum) => {
    return colIndices.map(col => {
      if (col === -1) return "";
      return String(mainData[rowNum-1][col] || "");
    });
  };

  const getLeavingRowCells = (rowNum) => {
    return colIndices.map(col => {
      if (col === -1) return "";
      const val = mainData[rowNum-1][col];
      if (val instanceof Date) {
        return val.getHours() + ":" + String(val.getMinutes()).padStart(2, '0');
      }
      return String(val || "");
    });
  };

  return {
    year: String(mainData[1][0]).replace(/年度/g, ""),
    startDate: (weekDates[0].getMonth() + 1) + "月" + weekDates[0].getDate() + "日",
    endDate: (weekDates[6].getMonth() + 1) + "月" + weekDates[6].getDate() + "日",
    weekNum: targetWeekNum,
    days: weekDates.map((d, idx) => `${d.getMonth() + 1}/${d.getDate()}(${dayLabels[idx]})`),
    holidayRow: getRowCells(7),
    dropdownSubjects: allSubjects,
    colorSubjects: colorSubjects,
    leavingTimes: leavingTimes,
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

  const saveSingleRow = (rowNum, rowDataArray) => {
    colIndices.forEach((colIdx, i) => {
      if (colIdx !== -1) mainSheet.getRange(rowNum, colIdx + 1).setValue(rowDataArray[i]);
    });
  };

  const saveSplitRow = (rowNum1, rowNum2, rowDataArray) => {
    colIndices.forEach((colIdx, i) => {
      if (colIdx !== -1) {
        mainSheet.getRange(rowNum1, colIdx + 1).setValue(rowDataArray[i].subject);
        mainSheet.getRange(rowNum2, colIdx + 1).setValue(rowDataArray[i].content);
      }
    });
  };

  try {
    // 【重要】11行目（行事行）は数式を保護するため、保存処理を完全にスキップします。
    saveSingleRow(12, clientData.morning);
    saveSplitRow(13, 14, clientData.p1);
    saveSplitRow(15, 16, clientData.p2);
    saveSingleRow(17, clientData.rest);
    saveSplitRow(18, 19, clientData.p3);
    saveSplitRow(20, 21, clientData.p4);
    saveSingleRow(22, clientData.lunch);
    saveSingleRow(23, clientData.clean);
    saveSingleRow(24, clientData.recess);
    saveSplitRow(25, 26, clientData.p5);
    saveSplitRow(27, 28, clientData.p6);
    saveSingleRow(29, clientData.leaving);
    saveSingleRow(30, clientData.homework);
    saveSingleRow(31, clientData.memo);

    return { success: true };
  } catch(e) {
    return { error: e.toString() };
  }
}

// Webアプリから行事・休日を設定シートの最後に追加する関数
function addNewConfigEventOrHoliday(type, dateStr, nameStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("設定");
    if (!configSheet) return { error: "設定シートが見つかりません。" };
    
    const targetDate = new Date(dateStr);
    if (isNaN(targetDate.getTime())) return { error: "日付の形式が正しくありません。" };
    
    if (type === 'event') {
      const fValues = configSheet.getRange("F1:F").getValues();
      let lastRowF = 1;
      for (let i = fValues.length - 1; i >= 0; i--) {
        if (fValues[i][0] !== "") { lastRowF = i + 1; break; }
      }
      const nextRow = Math.max(lastRowF + 1, 3);
      
      configSheet.getRange(nextRow, 6).setValue(targetDate);
      configSheet.getRange(nextRow, 7).setValue(nameStr);
      return { success: true, message: "行事を登録しました。" };
      
    } else if (type === 'holiday') {
      const hValues = configSheet.getRange("H1:H").getValues();
      let lastRowH = 1;
      for (let i = hValues.length - 1; i >= 0; i--) {
        if (hValues[i][0] !== "") { lastRowH = i + 1; break; }
      }
      const nextRow = Math.max(lastRowH + 1, 3);
      
      configSheet.getRange(nextRow, 8).setValue(targetDate);
      configSheet.getRange(nextRow, 9).setValue(nameStr);
      return { success: true, message: "休日を登録しました。" };
    }
    
    return { error: "不正な種別です。" };
  } catch(e) {
    return { error: e.toString() };
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