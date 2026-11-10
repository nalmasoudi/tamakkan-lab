/**
 * خادم تمكّن لاب — Google Apps Script
 * يستقبل: أحداث الاستخدام، ونتائج الاختبارات، وتقارير المتدربين، ويحفظها في Google Sheet.
 * ويعرض: إحصاءات مجمّعة فقط (بدون أسماء أو أرقام) للوحة الإحصاءات.
 * النشر: Deploy > New deployment > Web app (Execute as: Me — Who has access: Anyone)
 * بعد أي تعديل: Deploy > Manage deployments > Edit > New version (يبقى الرابط نفسه).
 */

/* معرّف ملف Google Sheet الذي تُحفظ فيه البيانات */
var SHEET_ID = '1rJO-od0EE4TBfxVGT1MYCNmhs95pWVYFcbwzDRqC78g';

var SHEETS = {
  events:   ['الوقت', 'الحدث', 'الجهاز', 'المستخدم', 'التجربة', 'المادة', 'عدد', 'من', 'اكتملت', 'الدرجة'],
  results:  ['الوقت', 'الاختبار', 'القناة', 'الرقم الأكاديمي', 'النتيجة ورمز التحقق'],
  messages: ['الوقت', 'القناة', 'الرسالة']
};

function sheet_(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); sh.setRightToLeft(true); }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(v, n) { return String(v == null ? '' : v).slice(0, n || 100); }
function num_(v) { v = Number(v); return isFinite(v) ? v : ''; }

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: 'bad json' }); }
  if (!body || e.postData.contents.length > 200000) return out_({ ok: false, err: 'too large' });
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var now = new Date();
    if (body.k === 'ev' && Array.isArray(body.list)) {
      var rows = body.list.slice(0, 100).map(function (o) {
        return [o.t ? new Date(Number(o.t)) : now, str_(o.ev, 12), str_(o.dv, 20), str_(o.u, 20), str_(o.lab, 40), str_(o.c, 40), num_(o.n), num_(o.of), num_(o.all), num_(o.sc)];
      });
      if (rows.length) { var sh = sheet_('events'); sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows); }
      return out_({ ok: true, n: rows.length });
    }
    if (body.k === 'res') { sheet_('results').appendRow([now, str_(body.e, 20), str_(body.q, 40), str_(body.i, 20), str_(body.sub, 60000)]); return out_({ ok: true }); }
    if (body.k === 'msg') { sheet_('messages').appendRow([now, str_(body.q, 40), str_(body.body, 60000)]); return out_({ ok: true }); }
    return out_({ ok: false, err: 'unknown' });
  } finally { lock.releaseLock(); }
}

/* إحصاءات مجمّعة فقط — لا تُرجع أي بيانات شخصية */
function doGet(e) {
  var cache = CacheService.getScriptCache(), hit = cache.get('stats');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var v = sheet_('events').getDataRange().getValues().slice(1);
  var dev = {}, usr = {}, labs = {}, done = {}, course = {}, days = {}, opens = 0, tasks = 0, exams = 0, visits = 0;
  v.forEach(function (r) {
    var t = new Date(r[0]), ev = r[1], d = r[2], u = r[3], lab = r[4], c = r[5], n = Number(r[6]) || 0;
    var day = Utilities.formatDate(t, 'Asia/Riyadh', 'yyyy-MM-dd');
    if (d) dev[d] = 1; if (u) usr[u] = 1;
    days[day] = days[day] || { u: {}, tasks: 0, opens: 0 };
    if (d) days[day].u[d] = 1;
    if (c) course[c] = course[c] || { opens: 0, tasks: 0, done: 0, users: {} };
    if (ev === 'visit') visits++;
    if (ev === 'open') { opens++; days[day].opens++; if (lab) labs[lab] = (labs[lab] || 0) + 1; if (c) { course[c].opens++; course[c].users[d] = 1; } }
    if (ev === 'task') { tasks += n; days[day].tasks += n; if (c) course[c].tasks += n; if (Number(r[8]) === 1) { done[d + '|' + lab] = 1; if (c) course[c].done++; } }
    if (ev === 'exam') exams++;
  });
  var o = {
    updated: new Date().toISOString(),
    devices: Object.keys(dev).length, users: Object.keys(usr).length, visits: visits,
    opens: opens, tasks: tasks, labsCompleted: Object.keys(done).length, exams: exams,
    courses: Object.keys(course).map(function (k) { var c = course[k]; return { id: k, opens: c.opens, tasks: c.tasks, done: c.done, users: Object.keys(c.users).length }; }),
    topLabs: Object.keys(labs).map(function (k) { return { id: k, opens: labs[k] }; }).sort(function (a, b) { return b.opens - a.opens; }).slice(0, 15),
    days: Object.keys(days).sort().slice(-90).map(function (k) { return { d: k, users: Object.keys(days[k].u).length, tasks: days[k].tasks, opens: days[k].opens }; })
  };
  var js = JSON.stringify(o); cache.put('stats', js, 300);
  return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JSON);
}
