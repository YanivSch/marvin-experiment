// ניסוי המבוך: קליטת קודי תוצאה לגיליון, וקריאתם ללוח הכיתה
// מדביקים ב-Apps Script של הגיליון ופורסים כ-Web app (Execute as: Me, Who has access: Anyone).
const SHEET_NAME = 'maze';

// MZ3-6 קטעים (אחד לכל שלב)-P ניקוד-כינוי. הפירוק המלא נעשה בלוח הכיתה (class.html).
const SOLO = '(\\d)\\.(\\d)\\.(\\d{1,2})\\.(\\d{1,2})';
const AI = '(\\d)\\.(\\d)\\.([NB])\\.([01])\\.(\\d{1,2})\\.(\\d{1,2})';
const RE = new RegExp('MZ3-' + [SOLO, AI, AI, SOLO, AI, AI].join('-') + '-P(\\d{1,3})-([^\\s]{1,12})');

function doPost(e) {
  const raw = (e.parameter && e.parameter.code) || (e.postData && e.postData.contents) || '';
  const m = String(raw).match(RE);
  if (!m) return reply('bad code');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // עמודות: זמן, קוד, ניקוד, כינוי
    SpreadsheetApp.getActive().getSheetByName(SHEET_NAME).appendRow([new Date(), m[0], Number(m[m.length - 2]), m[m.length - 1]]);
  } finally {
    lock.releaseLock();
  }
  return reply('ok');
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'list') return reply('alive');
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const last = sheet.getLastRow();
  const rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, 2).getValues()
    .filter(function (r) { return r[1]; })
    .map(function (r) { return { t: new Date(r[0]).getTime(), code: String(r[1]) }; });
  const json = JSON.stringify({ rows: rows });
  if (p.callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function reply(text) {
  return ContentService.createTextOutput(text);
}
