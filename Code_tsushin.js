
/**
 * 週案シートから指定された週のデータを取得し、縦型表示用に整形して返す
 */
function getWeekPlan(week) {
  const plans = getWeekPlans([week]);
  return plans[String(Number(week))] || {};
}

/** 複数週を1回のシート読込で返す。 */
function getWeekPlans(weeks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("週案");
  if (!sheet) return {};

  const lastRow = Math.min(sheet.getLastRow(), 35);
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 10 || lastColumn < 7) return {};
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const result = {};
  (weeks || []).forEach(function(week) {
    result[String(Number(week))] = buildWeekPlanFromValues_(week, values, lastColumn);
  });
  return result;
}

function buildWeekPlanFromValues_(week, values, lastColumn) {
  // 8行目(インデックス7)から週番号を検索
  let weekStartCol = -1;
  const targetWeek = Number(week); // 数値化して確実に比較

  for (let col = 0; col < lastColumn; col++) {
    const val = values[7][col];
    if (val !== "" && val !== null && !isNaN(val)) {
      if (Number(val) === targetWeek) {
        weekStartCol = col;
        break;
      }
    }
  }

  if (weekStartCol === -1) return {};

  const result = {
    dates: [],
    event: [],
    p1: [], p2: [], p3: [], p4: [], p5: [], p6: [],
    dismissal: [],
    homework: [],
    isHoliday: [],
    lunchDuty: "" 
  };

  // 給食当番の取得 (9行目：インデックス8)
  if (values[8] && values[8][weekStartCol]) {
    const dutyRaw = String(values[8][weekStartCol]).trim();
    if (dutyRaw === "A" || dutyRaw === "B") {
      result.lunchDuty = dutyRaw + "当番";
    } else {
      result.lunchDuty = dutyRaw;
    }
  }

  for (let i = 0; i < 7; i++) {
    const currentCol = weekStartCol + i;
    if (currentCol >= lastColumn) break;

    // 7行目(インデックス6)が「休」かどうか
    const isHoliday = (String(values[6][currentCol]).trim() === "休");
    result.isHoliday.push(isHoliday);

    // 日付整形
    let dateStr = "";
    const rawDate = values[9][currentCol]; // 10行目
    if (rawDate instanceof Date) {
      const m = rawDate.getMonth() + 1;
      const d = rawDate.getDate();
      const dayLabels = ["日", "月", "火", "水", "木", "金", "土"];
      const wStr = dayLabels[rawDate.getDay()];
      dateStr = m + "/" + d + "(" + wStr + ")";
    } else if (rawDate) {
      dateStr = String(rawDate);
    }
    result.dates.push(dateStr);

    // 行事予定（11行目）。12行目は朝の予定なので混在させない
    const eventValue = values[10][currentCol] ? String(values[10][currentCol]).trim() : "";
    result.event.push(eventValue);

    // 教科 (正しい行位置)
    result.p1.push(values[12][currentCol] ? String(values[12][currentCol]) : ""); // 13
    result.p2.push(values[14][currentCol] ? String(values[14][currentCol]) : ""); // 15
    result.p3.push(values[17][currentCol] ? String(values[17][currentCol]) : ""); // 18
    result.p4.push(values[19][currentCol] ? String(values[19][currentCol]) : ""); // 20
    result.p5.push(values[24][currentCol] ? String(values[24][currentCol]) : ""); // 25
    result.p6.push(values[26][currentCol] ? String(values[26][currentCol]) : ""); // 27
    
    // 下校時刻 (29行目)
    let dismissStr = "";
    const rawDismiss = values[28] ? values[28][currentCol] : ""; 
    if (rawDismiss instanceof Date) {
      const hours = String(rawDismiss.getHours()).padStart(2, '0');
      const minutes = String(rawDismiss.getMinutes()).padStart(2, '0');
      dismissStr = hours + ":" + minutes;
    } else if (rawDismiss) {
      dismissStr = String(rawDismiss);
    }
    result.dismissal.push(dismissStr);

    // 宿題 (30行目)
    let homeworkStr = "";
    const rawHomework = values[29] ? values[29][currentCol] : ""; 
    if (rawHomework) {
      homeworkStr = String(rawHomework).replace(/([^ \n])□/g, "$1\n□");
    }
    result.homework.push(homeworkStr);
  }
  // 設定ページで選択した項目だけを学級通信へ渡す。
  const fields = typeof getTsushinFieldSettings_ === 'function'
    ? getTsushinFieldSettings_()
    : { event: true, timetable: true, dismissal: true, homework: true, lunchDuty: true };
  if (!fields.event) result.event = result.event.map(function() { return ''; });
  if (!fields.timetable) ['p1','p2','p3','p4','p5','p6'].forEach(function(key) {
    result[key] = result[key].map(function() { return ''; });
  });
  if (!fields.dismissal) result.dismissal = result.dismissal.map(function() { return ''; });
  if (!fields.homework) result.homework = result.homework.map(function() { return ''; });
  if (!fields.lunchDuty) result.lunchDuty = '';
  return result;
}

/**
 * 【通信データ】シートから、対象の週やQRコード詳細設定を含むすべてのデータを読み込む関数
 */
function loadTsushinData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("通信データ");
  if (!sheet) return {};
  
  const v = sheet.getRange('B1:B20').getValues().map(function(row) { return row[0]; });
  return { issue:v[0], columnTitle:v[1], columnBody:v[2], notice:v[3], qrRentaku:v[4], qrHomework:v[5],
    upperType:v[6]||'table', upperColTitle:v[7]||'上段コラム', upperColBody:v[8]||'', lowerType:v[9]||'table',
    lowerColTitle:v[10]||'下段コラム', lowerColBody:v[11]||'', photoCaption:v[12]||'', upperWeek:v[13]||'11',
    lowerWeek:v[14]||'12', qrRenrakuLabel:v[15]||'連絡帳', qrRenrakuShow:v[16]!==false,
    qrHomeworkLabel:v[17]||'宿題チェック', qrHomeworkShow:v[18]!==false, issueDate:v[19]||'' };
}

function loadTsushinInitialData() {
  const form = loadTsushinData();
  const upperWeek = Number(form.upperWeek || 11);
  const lowerWeek = Number(form.lowerWeek || 12);
  return { form: form, plans: getWeekPlans([upperWeek, lowerWeek]) };
}

/**
 * アプリで入力されたすべての文章、レイアウト、対象週、QRコード表示設定を【通信データ】シートに保存する関数
 */
function saveTsushinData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("通信データ");
  if (!sheet) return "エラー：『通信データ』シートが見つかりません。";
  
  const values = [data.issue,data.columnTitle,data.columnBody,data.notice,data.qrRentaku,data.qrHomework,
    data.upperType,data.upperColTitle,data.upperColBody,data.lowerType,data.lowerColTitle,data.lowerColBody,
    data.photoCaption,data.upperWeek,data.lowerWeek,data.qrRenrakuLabel,data.qrRenrakuShow,
    data.qrHomeworkLabel,data.qrHomeworkShow,data.issueDate].map(function(value){return [value];});
  sheet.getRange('B1:B20').setValues(values);
  
  return "スプレッドシートに保存しました！";
}

//現在の入力データからPDFを生成し、各自のGoogleドライブに安全に保存する

function savePdfToDrive(formData, upperWeek) {
  try {
    // 1. 保存した当日の日付を取得 (例: 20260514)
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMdd");
    const year = now.getFullYear();
    
    // 2. 指定されたファイル名を組み立て
    const fileName = `${year}_今週のお知らせ_week${upperWeek}_${formattedDate}.pdf`;
    
    // 3. 印刷用HTMLテンプレート（WeeklyPrint.html）を読み込んでデータを流し込む
    const htmlOutput = HtmlService.createTemplateFromFile('WeeklyPrint');
    htmlOutput.data = formData;
    htmlOutput.upperWeek = upperWeek;
    
    const htmlContent = htmlOutput.evaluate().getContent();
    
    // 4. HTMLからBlobを作成
    const blob = Utilities.newBlob(htmlContent, "text/html", fileName).getAs("application/pdf");
    
    // 5. 【汎用化】コピーした先生自身のマイドライブ（ルートフォルダ）へ直接保存
    const file = DriveApp.createFile(blob);
    
    // 特定のフォルダーを指定させたい場合は、以下のように「設定シート」からIDを読み込むか、
    // 以下を有効化して、各自にIDを書き換えてもらってください。
    // const folderId = "ここに各自のフォルダーID"; 
    // const file = DriveApp.getFolderById(folderId).createFile(blob);
    
    return `Googleドライブに「${fileName}」としてPDFを保存しました！`;
  } catch (e) {
    return "PDF保存エラー: " + e.toString();
  }
}
