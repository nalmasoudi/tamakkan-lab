/**
 * خادم تمكّن لاب — Google Apps Script
 * يحفظ في Google Sheet: الحسابات، والشعب، وتقدّم المتدربين، والنتائج، والتقارير، وأحداث الاستخدام.
 * ويعرض للعامة: إحصاءات مجمّعة فقط (بدون أسماء أو أرقام).
 * النشر: Deploy > New deployment > Web app (Execute as: Me — Who has access: Anyone)
 * بعد أي تعديل: Deploy > Manage deployments > Edit > Version: New version (يبقى الرابط نفسه).
 */

/* معرّف ملف Google Sheet الذي تُحفظ فيه البيانات */
var SHEET_ID = '1rJO-od0EE4TBfxVGT1MYCNmhs95pWVYFcbwzDRqC78g';

var SHEETS = {
  events:   ['الوقت', 'الحدث', 'الجهاز', 'المستخدم', 'التجربة', 'المادة', 'عدد', 'من', 'اكتملت', 'الدرجة'],
  results:  ['الوقت', 'الاختبار', 'القناة', 'الرقم الأكاديمي', 'النتيجة ورمز التحقق'],
  messages: ['الوقت', 'القناة', 'الرسالة'],
  users:    ['الرقم التدريبي', 'الاسم', 'salt', 'hash', 'تاريخ التسجيل', 'آخر دخول', 'الشعب'],
  sections: ['رمز الشعبة', 'اسم الشعبة', 'المدرب', 'المواد', 'تاريخ الإنشاء', 'نشطة'],
  progress: ['الرقم التدريبي', 'التقدم', 'آخر تحديث'],
  staff:    ['المعرف', 'الاسم', 'الدور', 'salt', 'hash', 'تاريخ الإنشاء'],
  exams:    ['رمز الاختبار', 'بيانات الاختبار', 'المدرب', 'تاريخ الإنشاء']
};

/* ---------- أدوات عامة ---------- */
var SS_ = null;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(SHEET_ID)); }
function sheet_(name) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); sh.setRightToLeft(true); }
  return sh;
}
function rows_(name) { var v = sheet_(name).getDataRange().getValues(); v.shift(); return v; }
function findRow_(name, col, val) {
  var v = sheet_(name).getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][col]) === String(val)) return { i: i + 1, r: v[i] };
  return null;
}
function setRow_(name, i, r) { sheet_(name).getRange(i, 1, 1, r.length).setValues([r]); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(v, n) { return String(v == null ? '' : v).slice(0, n || 100); }
function num_(v) { v = Number(v); return isFinite(v) ? v : ''; }
function normId_(s) { return String(s || '').replace(/[٠-٩]/g, function (c) { return String(c.charCodeAt(0) - 1632); }).replace(/[۰-۹]/g, function (c) { return String(c.charCodeAt(0) - 1776); }).replace(/\s+/g, '').toLowerCase(); }
function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function rand_(n) { var a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = ''; for (var i = 0; i < n; i++) s += a.charAt(Math.floor(Math.random() * a.length)); return s; }
function hashPw_(pw, salt) {
  var h = salt + '|' + pw;
  for (var i = 0; i < 400; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  return h;
}
function secret_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function sign_(s) { return hex_(Utilities.computeHmacSha256Signature(s, secret_())).slice(0, 40); }
function mkTok_(role, id, days) {
  var p = Utilities.base64EncodeWebSafe(JSON.stringify({ r: role, i: id, x: Date.now() + days * 864e5 }));
  return p + '.' + sign_(p);
}
function readTok_(t) {
  try {
    var a = String(t || '').split('.'); if (a.length !== 2 || sign_(a[0]) !== a[1]) return null;
    var o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(a[0])).getDataAsString());
    return o.x > Date.now() ? o : null;
  } catch (e) { return null; }
}
/* منع التخمين: 8 محاولات خاطئة تقفل الحساب 10 دقائق */
function locked_(k) { return Number(CacheService.getScriptCache().get('f:' + k) || 0) >= 8; }
function fail_(k) { var c = CacheService.getScriptCache(), n = Number(c.get('f:' + k) || 0) + 1; c.put('f:' + k, String(n), 600); }
function okLogin_(k) { CacheService.getScriptCache().remove('f:' + k); }

/* ---------- الشعب ---------- */
function section_(code) {
  var f = findRow_('sections', 0, String(code || '').toUpperCase()); if (!f || String(f.r[5]) === '0') return null;
  return { code: f.r[0], name: f.r[1], trainer: f.r[2], courses: String(f.r[3] || '').split(',').filter(String) };
}
function sectionsOf_(list) { return String(list || '').split(',').filter(String).map(section_).filter(function (s) { return s; }); }
function staffName_(id) { var f = findRow_('staff', 0, id); return f ? f.r[1] : id; }
function publicSec_(s) { return { code: s.code, name: s.name, courses: s.courses, trainer: staffName_(s.trainer) }; }

/* ---------- المتدرب ---------- */
function profile_(f) {
  var p = findRow_('progress', 0, f.r[0]), prog = {};
  try { prog = p ? JSON.parse(p.r[1] || '{}') : {}; } catch (e) { prog = {}; }
  return { id: f.r[0], name: f.r[1], sections: sectionsOf_(f.r[6]).map(publicSec_), done: prog.done || {} };
}
function stRegister_(b) {
  var id = normId_(b.id), name = str_(b.name, 60).trim(), pw = String(b.pw || '');
  if (!/^[0-9a-z]{4,16}$/.test(id)) return { ok: false, err: 'الرقم التدريبي من 4 إلى 16 خانة (أرقام أو حروف إنجليزية).' };
  if (name.length < 3) return { ok: false, err: 'اكتب اسمك الثلاثي.' };
  if (pw.length < 6) return { ok: false, err: 'كلمة السر 6 خانات على الأقل.' };
  if (findRow_('users', 0, id)) return { ok: false, err: 'هذا الرقم مسجّل مسبقًا. سجّل دخولك، أو اطلب من مدرّبك إعادة تعيين كلمة السر.' };
  var sec = b.sec ? section_(b.sec) : null;
  if (b.sec && !sec) return { ok: false, err: 'رمز الشعبة غير صحيح.' };
  var salt = rand_(16), now = new Date();
  sheet_('users').appendRow([id, name, salt, hashPw_(pw, salt), now, now, sec ? sec.code : '']);
  return stLoginOk_(findRow_('users', 0, id));
}
function stLoginOk_(f) {
  var r = f.r.slice(); r[5] = new Date(); setRow_('users', f.i, r);
  return { ok: true, tok: mkTok_('st', f.r[0], 120), me: profile_({ i: f.i, r: r }) };
}
function stLogin_(b) {
  var id = normId_(b.id); if (locked_('st:' + id)) return { ok: false, err: 'محاولات كثيرة. حاول بعد 10 دقائق.' };
  var f = findRow_('users', 0, id);
  if (!f || hashPw_(String(b.pw || ''), f.r[2]) !== f.r[3]) { fail_('st:' + id); return { ok: false, err: f ? 'كلمة السر غير صحيحة.' : 'الرقم غير مسجّل. أنشئ حسابًا جديدًا.', nf: !f }; }
  okLogin_('st:' + id);
  if (b.sec) { var s = section_(b.sec); if (s) addSec_(f, s.code); f = findRow_('users', 0, id); }
  return stLoginOk_(f);
}
function addSec_(f, code) {
  var list = String(f.r[6] || '').split(',').filter(String);
  if (list.indexOf(code) < 0) { list.push(code); var r = f.r.slice(); r[6] = list.join(','); setRow_('users', f.i, r); }
}
function stUser_(b) { var t = readTok_(b.tok); if (!t || t.r !== 'st') return null; return findRow_('users', 0, t.i); }
function stMe_(b) { var f = stUser_(b); return f ? { ok: true, me: profile_(f) } : { ok: false, auth: true }; }
function stJoin_(b) {
  var f = stUser_(b); if (!f) return { ok: false, auth: true };
  var s = section_(b.sec); if (!s) return { ok: false, err: 'رمز الشعبة غير صحيح أو الشعبة مغلقة.' };
  addSec_(f, s.code); return { ok: true, me: profile_(findRow_('users', 0, f.r[0])) };
}
function stSync_(b) {
  var f = stUser_(b); if (!f) return { ok: false, auth: true };
  var p = findRow_('progress', 0, f.r[0]), cur = {};
  try { cur = p ? JSON.parse(p.r[1] || '{}') : {}; } catch (e) { cur = {}; }
  cur.done = cur.done || {};
  var inc = b.done || {}, n = 0;
  Object.keys(inc).slice(0, 300).forEach(function (k) { k = str_(k, 40); if (!cur.done[k]) { cur.done[k] = Number(inc[k]) || Date.now(); n++; } });
  var js = JSON.stringify(cur).slice(0, 45000);
  if (p) setRow_('progress', p.i, [f.r[0], js, new Date()]); else sheet_('progress').appendRow([f.r[0], js, new Date()]);
  return { ok: true, done: cur.done, added: n };
}

/* ---------- المدرّب والمسؤول ---------- */
function staffOf_(b) { var t = readTok_(b.tok); if (!t || (t.r !== 'tr' && t.r !== 'ad')) return null; var f = findRow_('staff', 0, t.i); return f ? { id: f.r[0], name: f.r[1], role: f.r[2], f: f } : null; }
function trState_() { return { ok: true, setup: rows_('staff').length === 0 }; }
function trSetup_(b) {
  if (rows_('staff').length) return { ok: false, err: 'تم إعداد المسؤول مسبقًا.' };
  var pw = String(b.pw || ''); if (pw.length < 8) return { ok: false, err: 'كلمة مرور المسؤول 8 خانات على الأقل.' };
  var salt = rand_(16); sheet_('staff').appendRow(['admin', str_(b.name || 'المسؤول', 60), 'admin', salt, hashPw_(pw, salt), new Date()]);
  return { ok: true, tok: mkTok_('ad', 'admin', 7), me: { id: 'admin', name: str_(b.name || 'المسؤول', 60), role: 'admin' } };
}
function trLogin_(b) {
  var id = normId_(b.id); if (locked_('tr:' + id)) return { ok: false, err: 'محاولات كثيرة. حاول بعد 10 دقائق.' };
  var f = findRow_('staff', 0, id);
  if (!f || hashPw_(String(b.pw || ''), f.r[3]) !== f.r[4]) { fail_('tr:' + id); return { ok: false, err: 'بيانات الدخول غير صحيحة.' }; }
  okLogin_('tr:' + id);
  return { ok: true, tok: mkTok_(f.r[2] === 'admin' ? 'ad' : 'tr', f.r[0], 7), me: { id: f.r[0], name: f.r[1], role: f.r[2] } };
}
function trSections_(b) {
  var s = staffOf_(b); if (!s) return { ok: false, auth: true };
  var users = rows_('users'), out = rows_('sections').filter(function (r) { return s.role === 'admin' || r[2] === s.id; }).map(function (r) {
    var n = users.filter(function (u) { return String(u[6] || '').split(',').indexOf(r[0]) >= 0; }).length;
    return { code: r[0], name: r[1], trainer: staffName_(r[2]), courses: String(r[3] || '').split(',').filter(String), active: String(r[5]) !== '0', n: n };
  });
  return { ok: true, sections: out, me: { id: s.id, name: s.name, role: s.role } };
}
function trSecSave_(b) {
  var s = staffOf_(b); if (!s) return { ok: false, auth: true };
  var name = str_(b.name, 60).trim(), courses = (b.courses || []).map(function (c) { return str_(c, 30); }).join(',');
  if (name.length < 2) return { ok: false, err: 'اكتب اسم الشعبة.' };
  if (b.code) {
    var f = findRow_('sections', 0, b.code); if (!f || (s.role !== 'admin' && f.r[2] !== s.id)) return { ok: false, err: 'لا تملك هذه الشعبة.' };
    var r = f.r.slice(); r[1] = name; r[3] = courses; if (b.active !== undefined) r[5] = b.active ? 1 : 0; setRow_('sections', f.i, r);
    return { ok: true, code: r[0] };
  }
  var code; do { code = rand_(6); } while (findRow_('sections', 0, code));
  sheet_('sections').appendRow([code, name, s.id, courses, new Date(), 1]);
  return { ok: true, code: code };
}
function trRoster_(b) {
  var s = staffOf_(b); if (!s) return { ok: false, auth: true };
  var f = findRow_('sections', 0, b.code); if (!f || (s.role !== 'admin' && f.r[2] !== s.id)) return { ok: false, err: 'لا تملك هذه الشعبة.' };
  var prog = {}; rows_('progress').forEach(function (p) { try { prog[p[0]] = JSON.parse(p[1] || '{}').done || {}; } catch (e) { } });
  var list = rows_('users').filter(function (u) { return String(u[6] || '').split(',').indexOf(f.r[0]) >= 0; }).map(function (u) {
    return { id: u[0], name: u[1], last: u[5] ? new Date(u[5]).getTime() : 0, done: prog[u[0]] || {} };
  });
  return { ok: true, section: { code: f.r[0], name: f.r[1], courses: String(f.r[3] || '').split(',').filter(String) }, students: list };
}
function trResetPw_(b) {
  var s = staffOf_(b); if (!s) return { ok: false, auth: true };
  var f = findRow_('users', 0, normId_(b.id)); if (!f) return { ok: false, err: 'الرقم غير موجود.' };
  if (s.role !== 'admin') {
    var mine = rows_('sections').filter(function (r) { return r[2] === s.id; }).map(function (r) { return r[0]; });
    if (!String(f.r[6] || '').split(',').some(function (c) { return mine.indexOf(c) >= 0; })) return { ok: false, err: 'المتدرب ليس في شعبك.' };
  }
  var pw = rand_(6).toLowerCase(), salt = rand_(16), r = f.r.slice(); r[2] = salt; r[3] = hashPw_(pw, salt); setRow_('users', f.i, r);
  okLogin_('st:' + f.r[0]);
  return { ok: true, pw: pw };
}
function adStaff_(b) {
  var s = staffOf_(b); if (!s || s.role !== 'admin') return { ok: false, auth: true };
  if (b.op === 'add') {
    var id = normId_(b.id), name = str_(b.name, 60).trim();
    if (!/^[0-9a-z._-]{3,20}$/.test(id)) return { ok: false, err: 'معرّف المدرب من 3 إلى 20 خانة إنجليزية أو أرقام.' };
    if (findRow_('staff', 0, id)) return { ok: false, err: 'المعرّف مستخدم.' };
    var pw = rand_(8).toLowerCase(), salt = rand_(16);
    sheet_('staff').appendRow([id, name || id, 'trainer', salt, hashPw_(pw, salt), new Date()]);
    return { ok: true, id: id, pw: pw };
  }
  if (b.op === 'reset') {
    var f = findRow_('staff', 0, normId_(b.id)); if (!f) return { ok: false, err: 'غير موجود.' };
    var p2 = rand_(8).toLowerCase(), s2 = rand_(16), r = f.r.slice(); r[3] = s2; r[4] = hashPw_(p2, s2); setRow_('staff', f.i, r);
    return { ok: true, id: f.r[0], pw: p2 };
  }
  return { ok: true, staff: rows_('staff').map(function (r) { return { id: r[0], name: r[1], role: r[2] }; }) };
}
function secInfo_(b) { var s = section_(b.sec); return s ? { ok: true, section: publicSec_(s) } : { ok: false, err: 'رمز الشعبة غير صحيح.' }; }

/* ---------- الاختبارات: روابط قصيرة ---------- */
function exSave_(b) {
  var s = staffOf_(b); if (!s) return { ok: false, auth: true };
  var P = b.P || {}; if (!P.e || !P.l || P.v !== 1) return { ok: false, err: 'بيانات الاختبار ناقصة.' };
  var code = str_(P.e, 12).toUpperCase(); if (findRow_('exams', 0, code)) return { ok: true, code: code };
  sheet_('exams').appendRow([code, JSON.stringify(P).slice(0, 45000), s.id, new Date()]);
  return { ok: true, code: code };
}
function exGet_(b) {
  var f = findRow_('exams', 0, str_(b.code, 12).toUpperCase()); if (!f) return { ok: false, err: 'رمز الاختبار غير موجود. تأكد من الرابط أو الرمز.' };
  try { return { ok: true, P: JSON.parse(f.r[1]) }; } catch (e) { return { ok: false, err: 'بيانات الاختبار تالفة.' }; }
}

/* ---------- نقاط الدخول ---------- */
var ACTIONS = { st_register: stRegister_, st_login: stLogin_, st_me: stMe_, st_join: stJoin_, st_sync: stSync_, sec_info: secInfo_,
  tr_state: trState_, tr_setup: trSetup_, tr_login: trLogin_, tr_sections: trSections_, tr_sec_save: trSecSave_, tr_roster: trRoster_, tr_resetpw: trResetPw_, ad_staff: adStaff_, ex_save: exSave_, ex_get: exGet_ };

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: 'bad json' }); }
  if (!body || e.postData.contents.length > 200000) return out_({ ok: false, err: 'too large' });
  if (body.a && ACTIONS[body.a]) {
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { return out_(ACTIONS[body.a](body)); } catch (x) { return out_({ ok: false, err: 'خطأ في الخادم' }); } finally { lock.releaseLock(); }
  }
  var lock2 = LockService.getScriptLock(); lock2.waitLock(10000);
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
  } finally { lock2.releaseLock(); }
}

/* إحصاءات مجمّعة فقط — لا تُرجع أي بيانات شخصية */
function doGet(e) {
  var cache = CacheService.getScriptCache(), hit = cache.get('stats');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var v = rows_('events');
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
    devices: Object.keys(dev).length, users: Object.keys(usr).length, accounts: rows_('users').length, sections: rows_('sections').length, visits: visits,
    opens: opens, tasks: tasks, labsCompleted: Object.keys(done).length, exams: exams,
    courses: Object.keys(course).map(function (k) { var c = course[k]; return { id: k, opens: c.opens, tasks: c.tasks, done: c.done, users: Object.keys(c.users).length }; }),
    topLabs: Object.keys(labs).map(function (k) { return { id: k, opens: labs[k] }; }).sort(function (a, b) { return b.opens - a.opens; }).slice(0, 15),
    days: Object.keys(days).sort().slice(-90).map(function (k) { return { d: k, users: Object.keys(days[k].u).length, tasks: days[k].tasks, opens: days[k].opens }; })
  };
  var js = JSON.stringify(o); cache.put('stats', js, 300);
  return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JSON);
}
