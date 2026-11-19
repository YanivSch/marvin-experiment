// ניסוי המבוך: קליטת קודי תוצאה לגיליון, וקריאתם ללוח הכיתה
// מדביקים ב-Apps Script של הגיליון ופורסים כ-Web app (Execute as: Me, Who has access: Anyone).
const SHEET_NAME = 'maze';

// MZ2-קטן(לבד.מילולי.הלך-כשטעה.סימון.הלך-כשטעה)-גדול(אותו דבר)-בונוס גבוה(סימון.הלך-כשטעה)
//    -7 זמנים בשניות-סדר(H/L)-P ניקוד-כינוי
const RE = /MZ2-(\d)\.(\d)\.(\d)\.(\d)\.(\d)-(\d)\.(\d)\.(\d)\.(\d)\.(\d)-(\d)\.(\d)-(\d{1,2})\.(\d{1,2})\.(\d{1,2})\.(\d{1,2})\.(\d{1,2})\.(\d{1,2})\.(\d{1,2})-([HL])-P(\d{1,2})-([^\s]{1,12})/;
const MAX = [2, 3, 1, 3, 1, 2, 3, 1, 3, 1, 2, 1];

function doPost(e) {
  const raw = (e.parameter && e.parameter.code) || (e.postData && e.postData.contents) || '';
  const m = String(raw).match(RE);
  if (!m) return reply('bad code');
  const v = m.slice(1, 20).map(Number);
  if (v.slice(0, 12).some(function (x, i) { return x > MAX[i]; })) return reply('bad code');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActive().getSheetByName(SHEET_NAME).appendRow([new Date(), m[0]].concat(v, [m[20], Number(m[21]), m[22]]));
  } finally {
    lock.releaseLock();
  }
  return reply('ok');
}

// ?action=list מחזיר את הזמן והקוד של כל שורה. עם callback=שם מוחזר JSONP.
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
