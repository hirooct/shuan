/**
 * Googleカレンダー → 週案の片方向同期。
 * 高度なCalendar APIは使わず、標準のCalendarAppだけで動作する。
 */
var SHUAN_CALENDAR_SYNC_CONFIG_KEY = 'CALENDAR_SYNC_CONFIG_V1';
var SHUAN_CALENDAR_SYNC_STATUS_KEY = 'CALENDAR_SYNC_STATUS_V1';
var SHUAN_CALENDAR_SYNC_SHEET = 'カレンダー同期';

function getCalendarSyncSetup() {
  const config = getCalendarSyncConfig_();
  const calendars = CalendarApp.getAllCalendars().map(function(calendar) {
    return { id: calendar.getId(), name: calendar.getName(), selected: calendar.isSelected() };
  }).sort(function(a, b) {
    if (a.selected !== b.selected) return a.selected ? -1 : 1;
    return a.name.localeCompare(b.name, 'ja');
  });
  return {
    calendars: calendars,
    config: config,
    status: getCalendarSyncStatus_()
  };
}

function saveCalendarSyncConfig(data) {
  const calendarId = String(data && data.calendarId || '').trim();
  if (!calendarId) return { error: '同期するGoogleカレンダーを選択してください。' };
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) return { error: '選択したGoogleカレンダーを読み込めません。' };

  const range = validateCalendarSyncRange_(data.startDate, data.endDate);
  if (range.error) return range;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = {
    calendarId: calendarId,
    calendarName: calendar.getName(),
    startDate: range.startText,
    endDate: range.endText,
    includeAllDay: !data || data.includeAllDay !== false,
    includeTimed: !data || data.includeTimed !== false,
    includeTimeLabel: Boolean(data && data.includeTimeLabel),
    autoSync: !data || data.autoSync !== false,
    spreadsheetId: ss.getId()
  };
  if (!config.includeAllDay && !config.includeTimed) {
    return { error: '終日予定または時刻付き予定のどちらかを選択してください。' };
  }
  PropertiesService.getDocumentProperties().setProperty(SHUAN_CALENDAR_SYNC_CONFIG_KEY, JSON.stringify(config));
  let triggerMessage = '';
  try {
    updateCalendarSyncTrigger_(config.autoSync);
    triggerMessage = config.autoSync ? '毎日5時台の自動同期も設定しました。' : '自動同期を解除しました。';
  } catch (e) {
    triggerMessage = '設定は保存しましたが、自動同期の設定に失敗しました: ' + e.message;
  }
  return { success: true, config: config, message: 'カレンダー同期設定を保存しました。' + triggerMessage };
}

function syncGoogleCalendarToWeeklyPlan() {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    const config = getCalendarSyncConfig_();
    if (!config.calendarId) return { error: '設定ページで同期するGoogleカレンダーを選択してください。' };
    const range = validateCalendarSyncRange_(config.startDate, config.endDate);
    if (range.error) return range;
    const calendar = CalendarApp.getCalendarById(config.calendarId);
    if (!calendar) return { error: '同期するGoogleカレンダーを読み込めません。権限と設定を確認してください。' };

    const desired = buildCalendarSyncEvents_(calendar, config, range.start, range.endExclusive);
    if (desired.length > 2000) return { error: '同期対象が2000件を超えています。期間を短くしてください。' };
    const result = writeCalendarSyncEvents_(desired, config);
    const status = {
      success: true,
      syncedAt: new Date().toISOString(),
      calendarName: calendar.getName(),
      added: result.added,
      updated: result.updated,
      removed: result.removed,
      total: desired.length
    };
    PropertiesService.getDocumentProperties().setProperty(SHUAN_CALENDAR_SYNC_STATUS_KEY, JSON.stringify(status));
    if (typeof clearShuanCaches_ === 'function') clearShuanCaches_();
    SpreadsheetApp.flush();
    return Object.assign({ message: 'Googleカレンダーを同期しました。追加' + result.added + '件、更新' + result.updated + '件、削除' + result.removed + '件。' }, status);
  } catch (e) {
    const status = { success: false, syncedAt: new Date().toISOString(), error: e.message };
    PropertiesService.getDocumentProperties().setProperty(SHUAN_CALENDAR_SYNC_STATUS_KEY, JSON.stringify(status));
    return { error: 'Googleカレンダーの同期に失敗しました: ' + e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function runScheduledCalendarSync() {
  const config = getCalendarSyncConfig_();
  if (config.autoSync !== false && config.calendarId) syncGoogleCalendarToWeeklyPlan();
}

function getCalendarSyncConfig_() {
  const props = PropertiesService.getDocumentProperties();
  try {
    const saved = JSON.parse(props.getProperty(SHUAN_CALENDAR_SYNC_CONFIG_KEY) || '{}');
    return Object.assign({ calendarId:'', calendarName:'', startDate:'', endDate:'', includeAllDay:true,
      includeTimed:true, includeTimeLabel:false, autoSync:true, spreadsheetId:'' }, saved);
  } catch (e) {
    return { calendarId:'', calendarName:'', startDate:'', endDate:'', includeAllDay:true,
      includeTimed:true, includeTimeLabel:false, autoSync:true, spreadsheetId:'' };
  }
}

function getCalendarSyncStatus_() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty(SHUAN_CALENDAR_SYNC_STATUS_KEY) || 'null'); }
  catch (e) { return null; }
}

function validateCalendarSyncRange_(startText, endText) {
  const parse = function(value) {
    const m = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const d = new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
    return d.getFullYear() === Number(m[1]) && d.getMonth() === Number(m[2])-1 && d.getDate() === Number(m[3]) ? d : null;
  };
  const start = parse(startText), end = parse(endText);
  if (!start || !end) return { error: '同期期間の開始日と終了日を正しく入力してください。' };
  if (end < start) return { error: '同期期間の終了日は開始日以降にしてください。' };
  if ((end.getTime()-start.getTime()) / 86400000 > 730) return { error: '同期期間は2年以内にしてください。' };
  const endExclusive = new Date(end); endExclusive.setDate(endExclusive.getDate()+1);
  const tz = Session.getScriptTimeZone();
  return { start:start, end:end, endExclusive:endExclusive,
    startText:Utilities.formatDate(start,tz,'yyyy-MM-dd'), endText:Utilities.formatDate(end,tz,'yyyy-MM-dd') };
}

function buildCalendarSyncEvents_(calendar, config, start, endExclusive) {
  const tz = Session.getScriptTimeZone();
  const output = [];
  calendar.getEvents(start, endExclusive).forEach(function(event) {
    const allDay = event.isAllDayEvent();
    if ((allDay && config.includeAllDay === false) || (!allDay && config.includeTimed === false)) return;
    const title = String(event.getTitle() || '').trim();
    if (!title) return;
    const eventId = event.getId();
    if (allDay) {
      const cursor = new Date(event.getStartTime()); cursor.setHours(0,0,0,0);
      const eventEnd = new Date(event.getEndTime()); eventEnd.setHours(0,0,0,0);
      while (cursor < eventEnd && cursor < endExclusive) {
        if (cursor >= start) {
          const dateText = Utilities.formatDate(cursor,tz,'yyyy-MM-dd');
          output.push({ key:config.calendarId+'|'+eventId+'|'+dateText, eventId:eventId, date:dateText, title:title });
        }
        cursor.setDate(cursor.getDate()+1);
      }
    } else {
      const eventStart = event.getStartTime();
      const dateText = Utilities.formatDate(eventStart,tz,'yyyy-MM-dd');
      const displayTitle = config.includeTimeLabel
        ? Utilities.formatDate(eventStart,tz,'HH:mm') + ' ' + title : title;
      output.push({ key:config.calendarId+'|'+eventId+'|'+dateText, eventId:eventId, date:dateText, title:displayTitle });
    }
  });
  const unique = {};
  output.forEach(function(item) { unique[item.key] = item; });
  return Object.keys(unique).map(function(key) { return unique[key]; }).sort(function(a,b) {
    return a.date.localeCompare(b.date) || a.title.localeCompare(b.title,'ja');
  });
}

function getCalendarSyncSpreadsheet_(config) {
  if (config && config.spreadsheetId) {
    try { return SpreadsheetApp.openById(config.spreadsheetId); } catch (e) { /* bound sheet fallback */ }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getCalendarSyncSheet_(ss) {
  let sheet = ss.getSheetByName(SHUAN_CALENDAR_SYNC_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SHUAN_CALENDAR_SYNC_SHEET);
    sheet.getRange(1,1,1,7).setValues([['同期キー','イベントID','日付','名称','設定行','カレンダーID','最終同期']]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function writeCalendarSyncEvents_(desired, config) {
  const ss = getCalendarSyncSpreadsheet_(config);
  const settings = ss.getSheetByName('設定');
  if (!settings) throw new Error('「設定」シートが見つかりません。');
  const syncSheet = getCalendarSyncSheet_(ss);
  const syncRows = syncSheet.getLastRow() > 1
    ? syncSheet.getRange(2,1,syncSheet.getLastRow()-1,7).getValues() : [];
  const oldByKey = {};
  syncRows.forEach(function(row) { if (row[0]) oldByKey[String(row[0])] = row; });

  const neededRows = Math.max(58, settings.getLastRow()-2, desired.length+20);
  if (settings.getMaxRows() < neededRows+2) settings.insertRowsAfter(settings.getMaxRows(), neededRows+2-settings.getMaxRows());
  const values = settings.getRange(3,6,neededRows,2).getValues();
  const occupied = values.map(function(row) { return row[0] !== '' || row[1] !== ''; });
  const desiredKeys = {};
  const nextSyncRows = [];
  let added=0, updated=0, removed=0;

  desired.forEach(function(item) {
    desiredKeys[item.key] = true;
    const old = oldByKey[item.key];
    let rowNumber = old ? Number(old[4]) : 0;
    let index = rowNumber - 3;
    const oldRowStillOwned = old && Number.isInteger(index) && index >= 0 && index < values.length &&
      ((formatCalendarSyncDate_(values[index][0]) === formatCalendarSyncDate_(old[2]) &&
        String(values[index][1] || '') === String(old[3] || '')) ||
       (values[index][0] === '' && values[index][1] === ''));
    if (!oldRowStillOwned) {
      index = occupied.findIndex(function(v) { return !v; });
      if (index < 0) throw new Error('行事の保存行が不足しています。同期期間を短くしてください。');
      rowNumber = index + 3;
      added++;
    }
    const date = parseCalendarSyncDate_(item.date);
    const oldDate = formatCalendarSyncDate_(values[index][0]);
    const oldTitle = String(values[index][1] || '');
    if (old && (oldDate !== item.date || oldTitle !== item.title)) updated++;
    values[index] = [date, item.title];
    occupied[index] = true;
    nextSyncRows.push([item.key,item.eventId,date,item.title,rowNumber,config.calendarId,new Date()]);
  });

  syncRows.forEach(function(row) {
    const key = String(row[0] || '');
    if (!key || desiredKeys[key]) return;
    const index = Number(row[4]) - 3;
    if (index >= 0 && index < values.length) {
      const sameDate = formatCalendarSyncDate_(values[index][0]) === formatCalendarSyncDate_(row[2]);
      const sameTitle = String(values[index][1] || '') === String(row[3] || '');
      if (sameDate && sameTitle) { values[index] = ['', '']; occupied[index] = false; removed++; }
    }
  });

  settings.getRange(3,6,neededRows,2).setValues(values);
  settings.getRange(3,6,neededRows,1).setNumberFormat('yyyy/MM/dd');
  if (syncSheet.getLastRow() > 1) syncSheet.getRange(2,1,syncSheet.getLastRow()-1,7).clearContent();
  if (nextSyncRows.length) syncSheet.getRange(2,1,nextSyncRows.length,7).setValues(nextSyncRows);
  return {added:added, updated:updated, removed:removed};
}

function parseCalendarSyncDate_(text) {
  const p = String(text).split('-');
  return new Date(Number(p[0]),Number(p[1])-1,Number(p[2]));
}

function formatCalendarSyncDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value,Session.getScriptTimeZone(),'yyyy-MM-dd');
  return String(value || '').replace(/\//g,'-');
}

function getCalendarSyncRowMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHUAN_CALENDAR_SYNC_SHEET);
  const result = {};
  if (!sheet || sheet.getLastRow() < 2) return result;
  sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues().forEach(function(row) {
    if (row[0] && row[4]) result[String(Number(row[4]))] = { calendarId:String(row[5]||''), key:String(row[0]) };
  });
  return result;
}

function updateCalendarSyncTrigger_(enabled) {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runScheduledCalendarSync') ScriptApp.deleteTrigger(trigger);
  });
  if (enabled) ScriptApp.newTrigger('runScheduledCalendarSync').timeBased().everyDays(1).atHour(5).create();
}
