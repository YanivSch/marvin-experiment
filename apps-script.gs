// הניסוי של מרווין וניסוי המבוך: קליטת קודי תוצאה לגיליון, וקריאתם ללוחות הכיתה
// מדביקים ב-Apps Script של הגיליון ופורסים כ-Web app (Execute as: Me, Who has access: Anyone).
const SHEET_NAME = 'results';
const MAZE_SHEET = 'maze'; // נוצרת אוטומטית בשליחה הראשונה

// ניסוי המבוך: MZ3-6 קטעים (אחד לכל שלב)-P ניקוד-כינוי. הפירוק המלא נעשה ב-maze/class.html.
const MZ_SOLO = '(\\d)\\.(\\d)\\.(\\d{1,2})\\.(\\d{1,2})';
const MZ_AI = '(\\d)\\.(\\d)\\.([NB])\\.([01])\\.(\\d{1,2})\\.(\\d{1,2})';
const MZ_RE = new RegExp('MZ3-' + [MZ_SOLO, MZ_AI, MZ_AI, MZ_SOLO, MZ_AI, MZ_AI].join('-') + '-P(\\d{1,3})-([^\\s]{1,12})');

function mazeSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(MAZE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MAZE_SHEET);
    sh.appendRow(['זמן', 'קוד', 'ניקוד', 'כינוי']);
  }
  return sh;
}
const GROUPS = { C: 'ביטחון בלבד', E: 'הסבר להמלצה', A: 'הסבר אדפטיבי' };

function doPost(e) {
  const raw = (e.parameter && e.parameter.code) ||
              (e.postData && e.postData.contents) || '';
  // גרסה 2 (ביקורות + גיוס): MRV2-קבוצה-לבד(0-6)-עם מרווין(0-14)-הסכים כשצדק(0-10)-הלך אחריו כשטעה(0-4)
  // גרסה 1 (ביקורות בלבד):  MRV-קבוצה-לבד(0-4)-עם מרווין(0-10)-הסכים כשצדק(0-7)-הלך אחריו כשטעה(0-3)
  const s = String(raw);

  const mz = s.match(MZ_RE);
  if (mz) {
    const lk = LockService.getScriptLock();
    lk.waitLock(10000);
    try {
      mazeSheet().appendRow([new Date(), mz[0], Number(mz[mz.length - 2]), mz[mz.length - 1]]);
    } finally {
      lk.releaseLock();
    }
    return reply('ok');
  }

  const v2 = s.match(/MRV2-([CEA])-(\d)-(\d{1,2})-(\d{1,2})-(\d)/);
  const m = v2 || s.match(/MRV-([CEA])-(\d)-(\d{1,2})-(\d)-(\d)/);
  if (!m) return reply('bad code');

  const max = v2 ? [6, 14, 10, 4] : [4, 10, 7, 3];
  const solo = Number(m[2]), team = Number(m[3]);
  const agreedRight = Number(m[4]), followedWrong = Number(m[5]);
  if (solo > max[0] || team > max[1] || agreedRight > max[2] || followedWrong > max[3]) {
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

// ?action=list מחזיר את הזמן והקוד של כל שורה, ללוח הכיתה. עם sheet=maze: של ניסוי המבוך.
// עם callback=שם מוחזר JSONP, כדי שעמוד באתר אחר יוכל לקרוא את התשובה.
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'list') return reply('alive');

  const sheet = p.sheet === 'maze' ? mazeSheet() : SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
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
