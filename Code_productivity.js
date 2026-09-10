/**
 * 日常入力支援・バックアップ機能（v3）
 * 既存シートのレイアウトを変えず、Document/User Properties に補助設定を保存する。
 */
var SHUAN_PRODUCTIVITY_KEYS = {
  PATTERNS: 'TIMETABLE_PATTERNS_V1',
  PHRASES: 'QUICK_PHRASES_V1',
  TSUSHIN: 'TSUSHIN_FIELDS_V1',
  BACKUPS: 'BACKUPS_V1',
  WIZARD: 'SETUP_WIZARD_DONE_V1'
};

function getProductivitySettings() {
  const props = PropertiesService.getDocumentProperties();
  return {
    patterns: parseJsonProperty_(props, SHUAN_PRODUCTIVITY_KEYS.PATTERNS, []),
    phrases: parseJsonProperty_(props, SHUAN_PRODUCTIVITY_KEYS.PHRASES, [
      { name: '基本の宿題', text: '□漢字\n□計算\n□音読' },
      { name: '持ち物', text: '□持ち物：' },
      { name: '委員会', text: '委員会活動' },
      { name: 'クラブ', text: 'クラブ活動' }
    ]),
    tsushinFields: parseJsonProperty_(props, SHUAN_PRODUCTIVITY_KEYS.TSUSHIN,
      { event: true, timetable: true, dismissal: true, homework: true, lunchDuty: true }),
    wizardDone: props.getProperty(SHUAN_PRODUCTIVITY_KEYS.WIZARD) === 'true'
  };
}

function saveProductivitySettings(data) {
  const props = PropertiesService.getDocumentProperties();
  const patterns = Array.isArray(data.patterns) ? data.patterns.slice(0, 30).map(normalizePattern_) : [];
  const phrases = Array.isArray(data.phrases) ? data.phrases.slice(0, 50).map(function(item) {
    return { name: String(item.name || '').trim(), text: String(item.text || '') };
  }).filter(function(item) { return item.name && item.text; }) : [];
  const fields = data.tsushinFields || {};
  props.setProperties({
    [SHUAN_PRODUCTIVITY_KEYS.PATTERNS]: JSON.stringify(patterns),
    [SHUAN_PRODUCTIVITY_KEYS.PHRASES]: JSON.stringify(phrases),
    [SHUAN_PRODUCTIVITY_KEYS.TSUSHIN]: JSON.stringify({
      event: fields.event !== false,
      timetable: fields.timetable !== false,
      dismissal: fields.dismissal !== false,
      homework: fields.homework !== false,
      lunchDuty: fields.lunchDuty !== false
    }),
    [SHUAN_PRODUCTIVITY_KEYS.WIZARD]: data.wizardDone ? 'true' : 'false'
  });
  return { success: true, message: '入力支援設定を保存しました。' };
}

function normalizePattern_(item) {
  const sourceDay = Math.max(0, Math.min(4, Number(item.sourceDay) || 0));
  const sourceStart = Math.max(1, Math.min(6, Number(item.sourceStart) || 1));
  const sourceEnd = Math.max(sourceStart, Math.min(6, Number(item.sourceEnd) || 6));
  const targetStart = Math.max(1, Math.min(6, Number(item.targetStart) || 1));
  return {
    name: String(item.name || '').trim().slice(0, 40), sourceDay: sourceDay,
    sourceStart: sourceStart, sourceEnd: sourceEnd, targetStart: targetStart,
    overwrite: item.overwrite !== false
  };
}

function parseJsonProperty_(props, key, fallback) {
  try { return JSON.parse(props.getProperty(key) || JSON.stringify(fallback)); }
  catch (e) { return fallback; }
}

function saveWeeklyDraft(weekNum, data) {
  const key = 'WEEKLY_DRAFT_' + String(weekNum);
  const payload = { savedAt: new Date().toISOString(), data: data };
  const json = JSON.stringify(payload);
  if (json.length > 80000) return { error: '下書きの容量が大きすぎます。通常保存してください。' };
  PropertiesService.getUserProperties().setProperty(key, json);
  return { success: true, savedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss') };
}

function getWeeklyDraft(weekNum) {
  return parseJsonProperty_(PropertiesService.getUserProperties(), 'WEEKLY_DRAFT_' + String(weekNum), null);
}

function deleteWeeklyDraft(weekNum) {
  PropertiesService.getUserProperties().deleteProperty('WEEKLY_DRAFT_' + String(weekNum));
  return { success: true };
}

function bulkAddConfigEntries(text) {
  const lines = String(text || '').split(/\r?\n/).map(function(v) { return v.trim(); }).filter(Boolean);
  if (!lines.length) return { error: '登録するデータを入力してください。' };
  if (lines.length > 300) return { error: '一度に登録できるのは300件までです。' };
  let added = 0, skipped = 0;
  const errors = [];
  lines.forEach(function(line, index) {
    const cols = line.split(/[\t,]/).map(function(v) { return v.trim(); });
    let type = 'event', date = '', name = '';
    if (/^(行事|event)$/i.test(cols[0])) { type = 'event'; date = cols[1]; name = cols.slice(2).join(' '); }
    else if (/^(休日|holiday)$/i.test(cols[0])) { type = 'holiday'; date = cols[1]; name = cols.slice(2).join(' '); }
    else { date = cols[0]; name = cols.slice(1).join(' '); }
    date = date.replace(/\//g, '-');
    const result = addNewConfigEventOrHoliday(type, date, name);
    if (result.success) added++;
    else if (String(result.error || '').includes('すでに登録')) skipped++;
    else errors.push((index + 1) + '行目: ' + (result.error || '登録失敗'));
  });
  return { success: errors.length === 0, added: added, skipped: skipped, errors: errors.slice(0, 10),
    message: added + '件登録、' + skipped + '件重複スキップ' + (errors.length ? '、' + errors.length + '件エラー' : '') };
}

function createAppBackup(label) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const settings = getAppDisplaySettings_();
  const name = '週案バックアップ_' + (settings.className || '未設定') + '_' + stamp + (label ? '_' + String(label).slice(0, 30) : '');
  const copied = DriveApp.getFileById(ss.getId()).makeCopy(name);
  const props = PropertiesService.getDocumentProperties();
  const backups = parseJsonProperty_(props, SHUAN_PRODUCTIVITY_KEYS.BACKUPS, []);
  const metadata = props.getProperties();
  delete metadata[SHUAN_PRODUCTIVITY_KEYS.BACKUPS];
  backups.unshift({ id: copied.getId(), name: copied.getName(), url: copied.getUrl(), createdAt: new Date().toISOString(), metadata: metadata });
  props.setProperty(SHUAN_PRODUCTIVITY_KEYS.BACKUPS, JSON.stringify(backups.slice(0, 20)));
  return { success: true, message: 'バックアップを作成しました。', backup: backups[0] };
}

function listAppBackups() {
  const items = parseJsonProperty_(PropertiesService.getDocumentProperties(), SHUAN_PRODUCTIVITY_KEYS.BACKUPS, []);
  return { backups: items.filter(function(item) { try { DriveApp.getFileById(item.id).getName(); return true; } catch(e) { return false; } }) };
}

function restoreAppBackup(fileId, confirmText) {
  if (String(confirmText || '') !== '復元する') return { error: '確認文字が一致しません。' };
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    createAppBackup('復元直前');
    const source = SpreadsheetApp.openById(String(fileId));
    const target = SpreadsheetApp.getActiveSpreadsheet();
    ['週案', '設定', '通信データ'].forEach(function(name) {
      const src = source.getSheetByName(name), dst = target.getSheetByName(name);
      if (!src || !dst) throw new Error('バックアップに「' + name + '」シートがありません。');
      const rows = src.getLastRow(), cols = src.getLastColumn();
      if (dst.getMaxRows() < rows) dst.insertRowsAfter(dst.getMaxRows(), rows - dst.getMaxRows());
      if (dst.getMaxColumns() < cols) dst.insertColumnsAfter(dst.getMaxColumns(), cols - dst.getMaxColumns());
      const values = src.getRange(1, 1, rows, cols).getValues();
      const formulas = src.getRange(1, 1, rows, cols).getFormulasR1C1();
      const restored = values.map(function(row, r) {
        return row.map(function(value, c) { return formulas[r][c] || value; });
      });
      const clearRows = Math.max(dst.getLastRow(), rows);
      const clearCols = Math.max(dst.getLastColumn(), cols);
      dst.getRange(1, 1, clearRows, clearCols).clearContent();
      dst.getRange(1, 1, rows, cols).setValues(restored);
    });
    const props = PropertiesService.getDocumentProperties();
    const backups = parseJsonProperty_(props, SHUAN_PRODUCTIVITY_KEYS.BACKUPS, []);
    const selected = backups.filter(function(item) { return item.id === String(fileId); })[0];
    if (selected && selected.metadata) props.setProperties(selected.metadata);
    clearShuanCaches_();
    SpreadsheetApp.flush();
    return { success: true, message: 'バックアップから復元しました。復元直前の状態も自動保存しています。' };
  } catch(e) { return { error: '復元に失敗しました: ' + e.message }; }
  finally { if (lock.hasLock()) lock.releaseLock(); }
}

function getTsushinFieldSettings_() {
  return getProductivitySettings().tsushinFields;
}
