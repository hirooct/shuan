/**
 * 設定シートのデータを取得して画面に渡す
 */
function getAppDisplaySettings_() {
  const props = PropertiesService.getDocumentProperties();
  return {
    appName: props.getProperty('APP_NAME') || '週案・学級通信',
    schoolName: props.getProperty('SCHOOL_NAME') || '',
    className: props.getProperty('CLASS_NAME') || '',
    teacherName: props.getProperty('TEACHER_NAME') || ''
  };
}

function getSettingsData() {
  const cache = CacheService.getDocumentCache();
  const cached = cache.get('SETTINGS_DATA_V1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
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
    
    const appSettings = getAppDisplaySettings_();
    const result = {
      appName: appSettings.appName,
      schoolName: appSettings.schoolName,
      className: appSettings.className,
      teacherName: appSettings.teacherName,
      nendo: nendo === null || nendo === undefined ? "" : String(nendo).trim(),
      nextNendo: nextNendo === null || nextNendo === undefined ? "" : String(nextNendo).trim(),
      slideUrl: slideUrl === null || slideUrl === undefined ? "" : String(slideUrl).trim(),
      vacation: vacation,
      subjects: subjects,
      dismissal: dismissal,
      timetable: timetable,
      colorSubjects: colorSubjects
    };
    
    cache.put('SETTINGS_DATA_V1', JSON.stringify(result), 600);
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
    const appName = String(data.appName || '').trim() || '週案・学級通信';
    PropertiesService.getDocumentProperties().setProperties({
      APP_NAME: appName,
      SCHOOL_NAME: String(data.schoolName || '').trim(),
      CLASS_NAME: String(data.className || '').trim(),
      TEACHER_NAME: String(data.teacherName || '').trim()
    });

    sheet.getRange("A2").setValue(data.nendo);
    sheet.getRange("C2").setValue(data.nextNendo);
    sheet.getRange("U13").setValue(data.slideUrl);
    
    sheet.getRange("M3:P3").setValues([data.vacationStart]);
    sheet.getRange("M4:P4").setValues([data.vacationEnd]);
    
    sheet.getRange("L8:L22").setValues(data.subjects.map(v => [v]));
    sheet.getRange("M8:M22").setValues(data.dismissal.map(v => [v]));
    sheet.getRange("V3:Z8").setValues(data.timetable);
    sheet.getRange("AB3:AB12").setValues(data.colorSubjects.map(v => [v]));
    clearShuanCaches_();
    return "設定を保存しました！";
  } catch (e) {
    return "保存エラー: " + e.toString();
  }
}

/** 設定画面の2種類の保存を1回の通信で完了する。 */
function saveAllSettingsData(data, productivityData) {
  const settingsResult = saveSettingsData(data);
  if (String(settingsResult).indexOf('エラー') === 0 || String(settingsResult).indexOf('保存エラー') === 0) {
    return { error: settingsResult };
  }
  const productivityResult = saveProductivitySettings(productivityData || {});
  if (productivityResult && productivityResult.error) return productivityResult;
  return { success: true, message: '設定を保存しました！' };
}

/**
 * 追加設定なしで利用できるCalendarAppから、日本の祝日を取得する
 */
function fetchAndWriteHolidays(currentYear, nextYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("設定");
  if (!sheet) return "エラー：「設定」シートが見つかりません。";

  const years = [Number(currentYear), Number(nextYear)];
  if (years.some(year => !Number.isInteger(year) || year < 2000 || year > 2100)) {
    return "エラー：正しい4桁の西暦を指定してください。";
  }

  try {
    const calendarId = 'ja.japanese.official#holiday@group.v.calendar.google.com';
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      return "エラー：日本の祝日カレンダーを取得できませんでした。";
    }

    const getHolidaysForYear = year => {
      const start = new Date(year, 0, 1);
      const end = new Date(year + 1, 0, 1);
      return calendar.getEvents(start, end)
        .filter(event => {
          const title = event.getTitle() || "";
          return !title.includes("銀行休業日") && !title.includes("大晦日");
        })
        .map(event => [
          Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "yyyy/MM/dd"),
          event.getTitle()
        ]);
    };

    const currentEvents = getHolidaysForYear(years[0]);
    const nextEvents = getHolidaysForYear(years[1]);

    sheet.getRange("A2").setValue(years[0]);
    sheet.getRange("C2").setValue(years[1]);
    sheet.getRange("A3:D60").clearContent();

    if (currentEvents.length) {
      sheet.getRange(3, 1, currentEvents.length, 2).setValues(currentEvents);
    }
    if (nextEvents.length) {
      sheet.getRange(3, 3, nextEvents.length, 2).setValues(nextEvents);
    }

    clearShuanCaches_();

    return "【完了】" + years[0] + "年と" + years[1] + "年の祝日を取得しました。";
  } catch (e) {
    return "祝日の取得に失敗しました: " + e.message;
  }
}
