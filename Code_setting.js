/**
 * 設定シートのデータを取得して画面に渡す
 */
function getSettingsData() {
  let currentStep = "開始";
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    currentStep = "「設定」シートの取得";
    const sheet = ss.getSheetByName("設定");
    if (!sheet) {
      throw new Error("「設定」シートが見つかりません。");
    }
    
    // どんなデータ形式でも安全に文字列にする補助関数
    const getSafeValues = (rangeStr) => {
      try {
        const range = sheet.getRange(rangeStr);
        const values = range.getValues();
        return values.map(row => 
          row.map(cell => {
            if (cell instanceof Date) {
              if (cell.getFullYear() < 1970) {
                return Utilities.formatDate(cell, "JST", "HH:mm");
              }
              return Utilities.formatDate(cell, "JST", "yyyy-MM-dd");
            }
            return cell === null || cell === undefined ? "" : String(cell).trim();
          })
        );
      } catch(e) {
        throw new Error(rangeStr + " のデータ取得に失敗: " + e.toString());
      }
    };

    currentStep = "A2, C2, U13セルの値取得";
    const nendo = sheet.getRange("A2").getValue();
    const nextNendo = sheet.getRange("C2").getValue();
    const slideUrl = sheet.getRange("U13").getValue();
    
    currentStep = "長期休暇(M2:P4)の取得";
    const vacation = getSafeValues("M2:P4");
    
    currentStep = "教科リスト(L8:L22)の取得";
    const subjects = getSafeValues("L8:L22").flat();
    
    currentStep = "下校時間(M8:M22)の取得";
    const dismissal = getSafeValues("M8:M22").flat();
    
    currentStep = "時間割(V3:Z8)の取得";
    const timetable = getSafeValues("V3:Z8");
    
    currentStep = "色変え教科(AB3:AB12)の取得";
    const colorSubjects = getSafeValues("AB3:AB12").flat();
    
    const result = {
      nendo: nendo === null || nendo === undefined ? "" : String(nendo).trim(),
      nextNendo: nextNendo === null || nextNendo === undefined ? "" : String(nextNendo).trim(),
      slideUrl: slideUrl === null || slideUrl === undefined ? "" : String(slideUrl).trim(),
      vacation: vacation,
      subjects: subjects,
      dismissal: dismissal,
      timetable: timetable,
      colorSubjects: colorSubjects
    };
    
    return result;
  } catch (e) {
    throw new Error(`[GASエラー発生位置: ${currentStep}] ${e.toString()}`);
  }
}

/**
 * 画面から送られたデータを設定シートに書き込む
 */
function saveSettingsData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("設定");
  if (!sheet) return "エラー：「設定」シートが見つかりません。";

  try {
    sheet.getRange("A2").setValue(data.nendo);
    sheet.getRange("C2").setValue(data.nextNendo);
    sheet.getRange("U13").setValue(data.slideUrl);
    
    sheet.getRange("M3:P3").setValues([data.vacationStart]);
    sheet.getRange("M4:P4").setValues([data.vacationEnd]);
    
    sheet.getRange("L8:L22").setValues(data.subjects.map(v => [v]));
    sheet.getRange("M8:M22").setValues(data.dismissal.map(v => [v]));
    sheet.getRange("V3:Z8").setValues(data.timetable);
    sheet.getRange("AB3:AB12").setValues(data.colorSubjects.map(v => [v]));
    
    return "設定を保存しました！";
  } catch (e) {
    return "保存エラー: " + e.toString();
  }
}

/**
 * GoogleカレンダーAPIを利用して指定年度と翌年度の祝日を書き込む
 */
function fetchAndWriteHolidays(currentYear, nextYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("設定");
  if (!sheet) return "「設定」シートが見つかりません。";

  if (!currentYear || isNaN(currentYear) || !nextYear || isNaN(nextYear)) {
    return "エラー：正しい4桁の西暦（年度）を指定してください。";
  }

  // 1. セル値を更新して、A3:D60をクリア
  sheet.getRange("A2").setValue(currentYear);
  sheet.getRange("C2").setValue(nextYear);
  sheet.getRange("A3:D60").clearContent();

  const calendarId = 'ja.japanese.official#holiday@group.v.calendar.google.com';

  // 指定した年の祝日を取得して配列で返す補助関数
  function getHolidaysForYear(year) {
    const timeMin = new Date(year + '-01-01T00:00:00Z').toISOString();
    const timeMax = new Date(year + '-12-31T23:59:59Z').toISOString();
    
    const response = Calendar.Events.list(calendarId, {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      orderBy: 'startTime'
    });

    if (!response.items || response.items.length === 0) return [];

    return response.items
      .filter(function(event) {
        const summary = event.summary || "";
        // 銀行休業日と大晦日は除外
        return !summary.includes("銀行休業日") && !summary.includes("大晦日");
      })
      .map(function(event) {
        // カレンダーの日付形式 YYYY-MM-DD を YYYY/MM/DD に変換
        const dateStr = event.start.date.replace(/-/g, "/");
        return [dateStr, event.summary];
      });
  }

  try {
    // 2. 指定年度（今年度）の祝日を取得して書き込み
    const currentEvents = getHolidaysForYear(currentYear);
    if (currentEvents.length > 0) {
      sheet.getRange(3, 1, currentEvents.length, 2).setValues(currentEvents);
    }

    // 3. 翌年度の祝日を取得して書き込み
    const nextEvents = getHolidaysForYear(nextYear);
    if (nextEvents.length > 0) {
      sheet.getRange(3, 3, nextEvents.length, 2).setValues(nextEvents);
    }

    return "【完了】Googleカレンダーから " + currentYear + "年と" + nextYear + "年の祝日を取得して書き込みました。";

  } catch (e) {
    return "カレンダーAPIからの取得に失敗しました。プロジェクトの設定で『Calendar API』が有効になっているか確認してください。エラー内容: " + e.toString();
  }
}