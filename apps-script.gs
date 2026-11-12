// הניסוי של מרווין: קליטת קודי תוצאה לגיליון, וקריאתם ללוח הכיתה
// מדביקים ב-Apps Script של הגיליון ופורסים כ-Web app (Execute as: Me, Who has access: Anyone).
const SHEET_NAME = 'results';
const GROUPS = { C: 'ביטחון בלבד', E: 'הסבר להמלצה', A: 'הסבר אדפטיבי' };

function doPost(e) {
  const raw = (e.parameter && e.parameter.code) ||
              (e.postData && e.postData.contents) || '';
  const m = String(raw).match(/MRV-([CEA])-(\d)-(\d{1,2})-(\d)-(\d)/);
  if (!m) return reply('bad code');

  const solo = Number(m[2]), team = Number(m[3]);
  const agreedRight = Number(m[4]), followedWrong = Number(m[5]);
  if (solo > 4 || team > 10 || agreedRight > 7 || followedWrong > 3) {
    return reply('bad code');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActive().getSheetByName(SHEET_NAME).appendRow([
      new Date(), m[0], GROUPS[m[1]], solo, team, agreedRight, followedWrong
    ]);
  } finally {
    lock.releaseLock();
  }
  return reply('ok');
}

// ?action=list מחזיר את הזמן והקוד של כל שורה, ללוח הכיתה.
// עם callback=שם מוחזר JSONP, כדי שעמוד באתר אחר יוכל לקרוא את התשובה.
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
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function reply(text) {
  return ContentService.createTextOutput(text);
}
