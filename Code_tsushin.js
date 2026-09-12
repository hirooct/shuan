
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
  
  const v = sheet.getRange('B1:B33').getValues().map(function(row) { return row[0]; });
  const result = { issue:v[0], columnTitle:v[1], columnBody:v[2], notice:v[3], qrRentaku:v[4], qrHomework:v[5],
    upperType:v[6]||'table', upperColTitle:v[7]||'上段コラム', upperColBody:v[8]||'', lowerType:v[9]||'table',
    lowerColTitle:v[10]||'下段コラム', lowerColBody:v[11]||'', photoCaption:v[12]||'', upperWeek:v[13]||'11',
    lowerWeek:v[14]||'12', qrRenrakuLabel:v[15]||'連絡帳', qrRenrakuShow:v[16]!==false,
    qrHomeworkLabel:v[17]||'宿題チェック', qrHomeworkShow:v[18]!==false, issueDate:v[19]||'' };
  // v3.2より前のシートでは、未保存だった画面初期値を不用意に空欄へしない。
  if (Number(v[32]) >= 2) {
    Object.assign(result, {mainTitle:v[20]||'', className:v[21]||'', upperTitle:v[22]||'', lowerTitle:v[23]||'',
      upperHasPhoto:v[24]===true, upperPhotoCaption:v[25]||'', showLunchDuty:v[26]!==false,
      theme:v[27]||'#1f3a5f,#e8eef7', photoTransform:parseTsushinJson_(v[28], {zoom:'1',x:'0',y:'0'}),
      upperPhotoTransform:parseTsushinJson_(v[29], {zoom:'1',x:'0',y:'0'}),
      photoFileId:v[30]||'', upperPhotoFileId:v[31]||''});
  }
  return result;
}

function loadTsushinInitialData() {
  const form = loadTsushinData();
  const upperWeek = Number(form.upperWeek || 11);
  const lowerWeek = Number(form.lowerWeek || 12);
  return { form: form, plans: getWeekPlans([upperWeek, lowerWeek]) };
}

/** 写真は初期画面と分けて遅延読込し、文字入力画面を先に表示する。 */
function loadTsushinPhotos() {
  const form = loadTsushinData();
  return {
    photoData: getTsushinPhotoData_(form.photoFileId),
    upperPhotoData: getTsushinPhotoData_(form.upperPhotoFileId)
  };
}

/**
 * アプリで入力されたすべての文章、レイアウト、対象週、QRコード表示設定を【通信データ】シートに保存する関数
 */
function saveTsushinData(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("通信データ");
    if (!sheet) throw new Error("『通信データ』シートが見つかりません。");

    const current = loadTsushinData();
    const photoFileId = updateTsushinPhoto_(data.photoData, current.photoFileId, data.removePhoto, data.issue, 'ひとこま');
    const upperPhotoFileId = updateTsushinPhoto_(data.upperPhotoData, current.upperPhotoFileId, data.removeUpperPhoto, data.issue, 'トピックス');
    const normalized = {
      issue:data.issue||'', columnTitle:data.columnTitle||'', columnBody:data.columnBody||'', notice:data.notice||'',
      qrRentaku:data.qrRentaku||'', qrHomework:data.qrHomework||'', upperType:data.upperType||'table',
      upperColTitle:data.upperColTitle||'', upperColBody:data.upperColBody||'', lowerType:data.lowerType||'table',
      lowerColTitle:data.lowerColTitle||'', lowerColBody:data.lowerColBody||'', photoCaption:data.photoCaption||'',
      upperWeek:data.upperWeek||'', lowerWeek:data.lowerWeek||'', qrRenrakuLabel:data.qrRenrakuLabel||'',
      qrRenrakuShow:data.qrRenrakuShow!==false, qrHomeworkLabel:data.qrHomeworkLabel||'',
      qrHomeworkShow:data.qrHomeworkShow!==false, issueDate:data.issueDate||'', mainTitle:data.mainTitle||'',
      className:data.className||'', upperTitle:data.upperTitle||'', lowerTitle:data.lowerTitle||'',
      upperHasPhoto:data.upperHasPhoto===true, upperPhotoCaption:data.upperPhotoCaption||'',
      showLunchDuty:data.showLunchDuty!==false, theme:data.theme||'#1f3a5f,#e8eef7',
      photoTransform:data.photoTransform||{zoom:'1',x:'0',y:'0'},
      upperPhotoTransform:data.upperPhotoTransform||{zoom:'1',x:'0',y:'0'},
      photoFileId:photoFileId, upperPhotoFileId:upperPhotoFileId
    };
    const values = [normalized.issue,normalized.columnTitle,normalized.columnBody,normalized.notice,
      normalized.qrRentaku,normalized.qrHomework,normalized.upperType,normalized.upperColTitle,
      normalized.upperColBody,normalized.lowerType,normalized.lowerColTitle,normalized.lowerColBody,
      normalized.photoCaption,normalized.upperWeek,normalized.lowerWeek,normalized.qrRenrakuLabel,
      normalized.qrRenrakuShow,normalized.qrHomeworkLabel,normalized.qrHomeworkShow,normalized.issueDate,
      normalized.mainTitle,normalized.className,normalized.upperTitle,normalized.lowerTitle,
      normalized.upperHasPhoto,normalized.upperPhotoCaption,normalized.showLunchDuty,normalized.theme,
      JSON.stringify(normalized.photoTransform),JSON.stringify(normalized.upperPhotoTransform),
      normalized.photoFileId,normalized.upperPhotoFileId,2].map(function(value){return [value];});
    sheet.getRange('B1:B33').setValues(values);
    SpreadsheetApp.flush();
    appendTsushinArchive_(normalized);
    return {ok:true, message:"学級通信を保存し、過去の保存にも追加しました。"};
  } finally {
    lock.releaseLock();
  }
}

function parseTsushinJson_(value, fallback) {
  if (!value) return fallback;
  try { return JSON.parse(String(value)); } catch (e) { return fallback; }
}

function getTsushinPhotoFolder_() {
  const props = PropertiesService.getDocumentProperties();
  const savedId = props.getProperty('TSUSHIN_PHOTO_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) { /* 再作成 */ }
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folderName = ss.getName() + '_学級通信写真';
  let parent = DriveApp.getRootFolder();
  try {
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) parent = parents.next();
  } catch (e) { /* 共有ドライブ等ではマイドライブ直下を使用 */ }
  const folder = parent.createFolder(folderName);
  props.setProperty('TSUSHIN_PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function updateTsushinPhoto_(dataUrl, currentId, remove, issue, label) {
  if (remove === true) return '';
  if (!dataUrl) return currentId || '';
  const match = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp|gif));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error(label + '写真の形式を確認してください。');
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 3 * 1024 * 1024) throw new Error(label + '写真が大きすぎます（3MB以下にしてください）。');
  const ext = match[1].split('/')[1].replace('jpeg','jpg');
  const safeIssue = String(issue || '号数未設定').replace(/[\\/:*?"<>|]/g, '_');
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const blob = Utilities.newBlob(bytes, match[1], safeIssue + '_' + label + '_' + stamp + '.' + ext);
  return getTsushinPhotoFolder_().createFile(blob).getId();
}

function getTsushinPhotoData_(fileId) {
  if (!fileId) return '';
  try {
    const blob = DriveApp.getFileById(String(fileId)).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return '';
  }
}

function getTsushinArchiveSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('学級通信アーカイブ');
  if (!sheet) {
    sheet = ss.insertSheet('学級通信アーカイブ');
    sheet.getRange(1,1,1,8).setValues([['ID','保存日時','号数','発行日','タイトル','上段週','下段週','保存データ']]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function appendTsushinArchive_(form) {
  const sheet = getTsushinArchiveSheet_();
  const id = Utilities.getUuid();
  const plans = getWeekPlans([Number(form.upperWeek), Number(form.lowerWeek)]);
  const payload = {version:2, form:form, plans:plans};
  sheet.appendRow([id,new Date(),form.issue||'',form.issueDate||'',form.mainTitle||'',form.upperWeek||'',form.lowerWeek||'',JSON.stringify(payload)]);
}

function listTsushinArchives(limit) {
  const sheet = getTsushinArchiveSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const count = Math.min(Math.max(Number(limit)||50,1),100,lastRow-1);
  const startRow = lastRow-count+1;
  return sheet.getRange(startRow,1,count,7).getValues().reverse().map(function(row) {
    return {id:String(row[0]), savedAt:row[1] instanceof Date ? Utilities.formatDate(row[1],'Asia/Tokyo','yyyy/MM/dd HH:mm') : String(row[1]||''),
      issue:String(row[2]||''), issueDate:String(row[3]||''), title:String(row[4]||''),
      upperWeek:String(row[5]||''), lowerWeek:String(row[6]||'')};
  });
}

function loadTsushinArchive(archiveId) {
  const sheet = getTsushinArchiveSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('保存履歴がありません。');
  const values = sheet.getRange(2,1,lastRow-1,8).getValues();
  for (let i=values.length-1; i>=0; i--) {
    if (String(values[i][0]) === String(archiveId)) {
      const saved = parseTsushinJson_(values[i][7], null);
      if (!saved) throw new Error('保存データを読み込めませんでした。');
      const form = saved.form || saved;
      return {form:form, plans:saved.plans||null,
        photos:{photoData:getTsushinPhotoData_(form.photoFileId), upperPhotoData:getTsushinPhotoData_(form.upperPhotoFileId)}};
    }
  }
  throw new Error('指定した保存履歴が見つかりません。');
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
