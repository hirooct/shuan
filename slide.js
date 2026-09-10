/**
 * 修正版：学級スライド自動生成処理
 * モードに応じて、日付・宿題の列と時間割・備考の列を適切に選択します。
 */
function executeSlideGeneration(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  if (!configSheet) return "エラー：「設定」シートが見つかりません。";
  
  const u13Value = configSheet.getRange("U13").getValue().toString().trim();
  const slideIdMatch = u13Value.match(/\/d\/([a-zA-Z0-9-_]+)/);
  const slideId = slideIdMatch ? slideIdMatch[1] : u13Value;
  if (!slideId) return "エラー：スライドURLまたはIDが正しくありません。";

  const sheet = ss.getSheetByName("週案"); 
  const lastColumn = sheet.getLastColumn();
  const dateRow = 10;
  const statusRow = 7;
  
  const dates = sheet.getRange(dateRow, 1, 1, lastColumn).getValues()[0];
  const statuses = sheet.getRange(statusRow, 1, 1, lastColumn).getValues()[0];

  const now = new Date();
  now.setHours(0, 0, 0, 0);

  // 💡 【重要】今日以降の「出」の列を順番に3つ取得する
  let schoolDayColumns = [];
  for (let i = 0; i < dates.length; i++) {
    let checkDate = new Date(dates[i]);
    if (dates[i] instanceof Date) {
      checkDate.setHours(0, 0, 0, 0);
      // 今日以降、かつステータスが「出」の列をピックアップ
      if (checkDate >= now && String(statuses[i]).trim() === "出") {
        schoolDayColumns.push(i + 1);
        if (schoolDayColumns.length >= 3) break; // 3日分確保すれば十分
      }
    }
  }

  if (schoolDayColumns.length < 2) {
    return "エラー：週案シートに十分な授業日（出）が見つかりませんでした。";
  }

  let homeworkCol, scheduleCol;

  // 実行日が「出」の日かどうかを判定
  const firstSchoolDay = new Date(dates[schoolDayColumns[0] - 1]);
  firstSchoolDay.setHours(0, 0, 0, 0);
  const isTodaySchoolDay = firstSchoolDay.getTime() === now.getTime();

  if (mode === 'next' && isTodaySchoolDay && schoolDayColumns.length < 3) {
    return "エラー：次の授業日の時間割を作るための授業日が不足しています。";
  }

  if (mode === 'next') {
    // ⏰ 【次の日モード】
    // 6/23(日)に実行した場合：schoolDayColumns[0]は6/24、[1]は6/25
    homeworkCol = schoolDayColumns[0]; // 宿題・日付は 6/24
    scheduleCol = schoolDayColumns[1]; // 時間割・備考は 6/25
    
    // もし今日が「出」の日(月曜)で、火曜のスライドを作りたい場合
    if (isTodaySchoolDay) {
      homeworkCol = schoolDayColumns[1]; // 宿題・日付は 6/25
      scheduleCol = schoolDayColumns[2]; // 時間割・備考は 6/26
    }
  } else {
    // ☀️ 【当日モード】
    homeworkCol = schoolDayColumns[0]; 
    scheduleCol = schoolDayColumns[1];
  }

  // データ抽出
  const targetDate = new Date(dates[homeworkCol - 1]);
  const dateStringTitle = Utilities.formatDate(targetDate, "JST", "M月d日") + "（" + getDayOfWeek(targetDate) + "）";

  const firstDataCol = Math.min(homeworkCol, scheduleCol);
  const block = sheet.getRange(13, firstDataCol, 19, Math.abs(scheduleCol - homeworkCol) + 1).getValues();
  const scheduleOffset = scheduleCol - firstDataCol;
  const homeworkOffset = homeworkCol - firstDataCol;
  const scheduleData = {
    date: dateStringTitle,
    p1: block[0][scheduleOffset],
    p2: block[2][scheduleOffset],
    p3: block[5][scheduleOffset],
    p4: block[7][scheduleOffset],
    p5: block[12][scheduleOffset],
    p6: block[14][scheduleOffset],
    homework: block[17][homeworkOffset],
    bikou: block[18][scheduleOffset] || ""
  };

  try {
    const presentation = SlidesApp.openById(slideId);
    const slides = presentation.getSlides();
    const lastSlide = slides[slides.length - 1]; 
    const newSlide = lastSlide.duplicate(); 
    newSlide.move(0); 

    newSlide.replaceAllText("{{日付}}", scheduleData.date);
    newSlide.replaceAllText("{{1}}", scheduleData.p1);
    newSlide.replaceAllText("{{2}}", scheduleData.p2);
    newSlide.replaceAllText("{{3}}", scheduleData.p3);
    newSlide.replaceAllText("{{4}}", scheduleData.p4);
    newSlide.replaceAllText("{{5}}", scheduleData.p5);
    newSlide.replaceAllText("{{6}}", scheduleData.p6);
    newSlide.replaceAllText("{{宿題}}", scheduleData.homework);
    newSlide.replaceAllText("{{備考}}", scheduleData.bikou); 
    
    return `【完了】\n${dateStringTitle} のスライドを作成しました！\n（宿題参照:${homeworkCol}列 / 時間割参照:${scheduleCol}列）`;
  } catch (e) {
    return "スライド作成エラー: " + e.message;
  }
}

function getDayOfWeek(date) {
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
  return weekdays[date.getDay()];
}
