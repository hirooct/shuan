/**
 * 配付・初期設定・保守機能
 */
var SHUAN_APP_VERSION = '3.0.0';
var SHUAN_APP_UPDATED_AT = '2026-09-09';
var SHUAN_RESET_CONFIRM_TEXT = '初期化する';

function getAppVersionInfo() {
  return {
    version: SHUAN_APP_VERSION,
    updatedAt: SHUAN_APP_UPDATED_AT
  };
}

function getSystemStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['はじめに', '週案', '設定', '通信データ'];
  const checks = [];

  requiredSheets.forEach(name => {
    checks.push({
      key: 'sheet_' + name,
      label: '「' + name + '」シート',
      status: ss.getSheetByName(name) ? 'ok' : 'error',
      message: ss.getSheetByName(name) ? '確認できました。' : 'シートが見つかりません。'
    });
  });

  const settings = ss.getSheetByName('設定');
  const weekly = ss.getSheetByName('週案');
  const appSettings = getAppDisplaySettings_();

  checks.push({
    key: 'identity',
    label: '学校・学級情報',
    status: appSettings.schoolName && appSettings.className ? 'ok' : 'warning',
    message: appSettings.schoolName && appSettings.className
      ? appSettings.schoolName + ' / ' + appSettings.className
      : '学校名または学年・組が未設定です。'
  });

  if (settings) {
    const currentYear = Number(settings.getRange('A2').getValue());
    const nextYear = Number(settings.getRange('C2').getValue());
    const validYears = Number.isInteger(currentYear) && Number.isInteger(nextYear) &&
      currentYear >= 2000 && nextYear === currentYear + 1;
    checks.push({
      key: 'years',
      label: '年度設定',
      status: validYears ? 'ok' : 'error',
      message: validYears
        ? currentYear + '年 / ' + nextYear + '年'
        : '指定年度と翌年度を正しく設定してください。'
    });

    const weekValues = settings.getRange('O9:P60').getValues();
    const validWeekCount = weekValues.filter(row =>
      row[0] !== '' && row[1] instanceof Date && !isNaN(row[1].getTime())
    ).length;
    checks.push({
      key: 'weeks',
      label: '週番号・開始日',
      status: validWeekCount >= 50 ? 'ok' : 'error',
      message: validWeekCount + '週分を確認しました。'
    });

    const timetableCount = settings.getRange('V3:Z8').getDisplayValues()
      .flat().filter(Boolean).length;
    checks.push({
      key: 'timetable',
      label: '基本時間割',
      status: timetableCount ? 'ok' : 'warning',
      message: timetableCount ? timetableCount + '件設定されています。' : '基本時間割が未設定です。'
    });

    const slideUrl = String(settings.getRange('U13').getValue() || '').trim();
    checks.push({
      key: 'slide',
      label: '連絡帳用スライド',
      status: slideUrl ? 'ok' : 'warning',
      message: slideUrl ? '設定されています。' : 'スライド機能を使う場合は設定してください。'
    });
  }

  if (weekly) {
    const lastColumn = weekly.getLastColumn();
    const dates = lastColumn >= 2
      ? weekly.getRange(10, 2, 1, lastColumn - 1).getValues()[0]
      : [];
    const dateCount = dates.filter(value => value instanceof Date && !isNaN(value.getTime())).length;
    checks.push({
      key: 'weekly_dates',
      label: '週案の日付',
      status: dateCount >= 300 ? 'ok' : (dateCount ? 'warning' : 'error'),
      message: dateCount + '日分の日付を確認しました。'
    });

    const eventFormulas = lastColumn >= 2
      ? weekly.getRange(11, 2, 1, lastColumn - 1).getFormulas()[0].filter(Boolean).length
      : 0;
    checks.push({
      key: 'event_formulas',
      label: '行事表示の数式',
      status: eventFormulas ? 'ok' : 'warning',
      message: eventFormulas
        ? eventFormulas + '列で数式を確認しました。'
        : '行事行の数式を確認できませんでした。'
    });
  }

  const webAppUrl = ScriptApp.getService().getUrl();
  checks.push({
    key: 'deployment',
    label: 'ウェブアプリ',
    status: webAppUrl ? 'ok' : 'warning',
    message: webAppUrl ? 'デプロイURLを確認しました。' : 'ウェブアプリが未デプロイです。'
  });

  return {
    version: SHUAN_APP_VERSION,
    updatedAt: SHUAN_APP_UPDATED_AT,
    ready: !checks.some(check => check.status === 'error'),
    checks: checks
  };
}

function getConfigEntries() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  if (!sheet) return { error: '「設定」シートが見つかりません。' };

  const lastRow = Math.max(sheet.getLastRow(), 3);
  const count = lastRow - 2;
  const timeZone = Session.getScriptTimeZone();
  const result = [];

  [
    { type: 'event', label: '行事', dateCol: 6 },
    { type: 'holiday', label: '休日', dateCol: 8 }
  ].forEach(config => {
    const values = sheet.getRange(3, config.dateCol, count, 2).getValues();
    values.forEach((row, index) => {
      if (row[0] === '' && row[1] === '') return;
      let date = '';
      if (row[0] instanceof Date && !isNaN(row[0].getTime())) {
        date = Utilities.formatDate(row[0], timeZone, 'yyyy-MM-dd');
      } else {
        date = String(row[0] || '');
      }
      result.push({
        type: config.type,
        typeLabel: config.label,
        row: index + 3,
        date: date,
        name: String(row[1] || '')
      });
    });
  });

  result.sort((a, b) => a.date.localeCompare(b.date) || a.type.localeCompare(b.type));
  return { entries: result };
}

function updateConfigEntry(type, rowNumber, dateStr, nameStr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  if (!sheet) return { error: '「設定」シートが見つかりません。' };

  const row = Number(rowNumber);
  const dateCol = type === 'event' ? 6 : type === 'holiday' ? 8 : 0;
  if (!dateCol || !Number.isInteger(row) || row < 3 || row > sheet.getMaxRows()) {
    return { error: '更新対象が正しくありません。' };
  }

  const name = String(nameStr || '').trim();
  const parts = String(dateStr || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!name || !parts) return { error: '日付と名称を正しく入力してください。' };

  const date = new Date(Number(parts[1]), Number(parts[2]) - 1, Number(parts[3]));
  if (date.getFullYear() !== Number(parts[1]) ||
      date.getMonth() !== Number(parts[2]) - 1 ||
      date.getDate() !== Number(parts[3])) {
    return { error: '存在しない日付です。' };
  }

  sheet.getRange(row, dateCol, 1, 2).setValues([[date, name]]);
  sheet.getRange(row, dateCol).setNumberFormat('yyyy/MM/dd');
  SpreadsheetApp.flush();
  return { success: true, message: '更新しました。' };
}

function deleteConfigEntry(type, rowNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  if (!sheet) return { error: '「設定」シートが見つかりません。' };

  const row = Number(rowNumber);
  const dateCol = type === 'event' ? 6 : type === 'holiday' ? 8 : 0;
  if (!dateCol || !Number.isInteger(row) || row < 3 || row > sheet.getMaxRows()) {
    return { error: '削除対象が正しくありません。' };
  }

  sheet.getRange(row, dateCol, 1, 2).clearContent();
  SpreadsheetApp.flush();
  return { success: true, message: '削除しました。' };
}

function copySlideTemplate(sourceUrl) {
  const url = String(sourceUrl || '').trim();
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  const fileId = match ? match[1] : (/^[a-zA-Z0-9-_]+$/.test(url) ? url : '');
  if (!fileId) return { error: 'GoogleスライドのURLまたはIDを入力してください。' };

  try {
    const source = DriveApp.getFileById(fileId);
    if (source.getMimeType() !== MimeType.GOOGLE_SLIDES) {
      return { error: '指定されたファイルはGoogleスライドではありません。' };
    }

    const settings = getAppDisplaySettings_();
    const suffix = [settings.schoolName, settings.className].filter(Boolean).join('_');
    const copyName = source.getName() + (suffix ? '_' + suffix : '_コピー');
    const copied = source.makeCopy(copyName);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
    if (!sheet) return { error: '「設定」シートが見つかりません。' };
    sheet.getRange('U13').setValue(copied.getUrl());

    return {
      success: true,
      message: 'スライドを自分のドライブにコピーし、URLを設定しました。',
      url: copied.getUrl()
    };
  } catch (e) {
    return { error: 'スライドのコピーに失敗しました: ' + e.message };
  }
}

function initializeForDistribution(confirmText) {
  if (String(confirmText || '').trim() !== SHUAN_RESET_CONFIRM_TEXT) {
    return { error: '確認文字が一致しないため、初期化しませんでした。' };
  }

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const weekly = ss.getSheetByName('週案');
    const settings = ss.getSheetByName('設定');
    const communication = ss.getSheetByName('通信データ');

    const missing = [];
    if (!weekly) missing.push('週案');
    if (!settings) missing.push('設定');
    if (!communication) missing.push('通信データ');
    if (missing.length) {
      return { error: '必要なシートがありません: ' + missing.join('、') };
    }

    // 初期化に進む直前の完全コピーを必ず残す。
    createAppBackup('初期化直前');

    const lastColumn = weekly.getLastColumn();
    if (lastColumn >= 2) {
      weekly.getRange(9, 2, 1, lastColumn - 1).clearContent();
      weekly.getRange(12, 2, 21, lastColumn - 1).clearContent();
    }

    settings.getRange('A2:D60').clearContent();
    settings.getRange(2, 6, settings.getMaxRows() - 1, 4).clearContent();
    settings.getRange('M3:P4').clearContent();
    settings.getRange('L8:M22').clearContent();
    settings.getRange('U13').clearContent();
    settings.getRange('V3:Z8').clearContent();
    settings.getRange('AB3:AB12').clearContent();

    communication.getRange(1, 2, Math.max(communication.getLastRow(), 20), 1).clearContent();
    const documentProps = PropertiesService.getDocumentProperties();
    ['APP_NAME', 'SCHOOL_NAME', 'CLASS_NAME', 'TEACHER_NAME',
      SHUAN_PRODUCTIVITY_KEYS.PATTERNS, SHUAN_PRODUCTIVITY_KEYS.PHRASES,
      SHUAN_PRODUCTIVITY_KEYS.TSUSHIN, SHUAN_PRODUCTIVITY_KEYS.WIZARD
    ].forEach(function(key) { documentProps.deleteProperty(key); });
    SpreadsheetApp.flush();

    return {
      success: true,
      message: '入力データを初期化しました。シート構造と週案の数式は残しています。'
    };
  } catch (e) {
    return { error: '初期化に失敗しました: ' + e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
