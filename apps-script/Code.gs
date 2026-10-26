/**
 * مطبخنا النباتي — API للشيت "plant based diet"
 * 1) Extensions > Apps Script > الصق الكود ده > Save
 * 2) شغّل الدالة setup مرة واحدة (هتطلب صلاحيات — وافق)
 * 3) Deploy > New deployment > Web app > Execute as: Me / Who has access: Anyone
 * 4) انسخ رابط /exec وحطه في config.js في الريبو
 */

const IMAGES_FOLDER = 'meals_Images';
const SH = { MEALS: 'meals', PLAN: 'Plan', SHOP: 'Shopping', SET: 'Settings', LOG: 'SwapLog' };
const PLAN_HEAD = ['week_start', 'date', 'day', 'slot', 'meal_id', 'meal_ar', 'kcal', 'status', 'locked', 'updated_at'];
const META_HEAD = ['photo_url', 'contents_urls', 'id', 'ar_name', 'slots', 'group', 'kcal', 'protein_g', 'prep_min', 'serves', 'ingredients'];
// الأعمدة الجديدة تبدأ من G (العمود 7) — A:F زي ما هي
const META_COL = 7;

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, head) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    if (head) { sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1); }
  }
  return sh;
}

/* ============ تشغيل مرة واحدة ============ */
function setup() {
  sheet_(SH.PLAN, PLAN_HEAD);
  sheet_(SH.SHOP, ['scope', 'item', 'checked', 'updated_at']);
  sheet_(SH.SET, ['key', 'value']);
  sheet_(SH.LOG, ['timestamp', 'date', 'slot', 'from', 'to', 'changes']);
  const m = ss_().getSheetByName(SH.MEALS);
  m.getRange(1, META_COL, 1, META_HEAD.length).setValues([META_HEAD]).setFontWeight('bold');
  setupImages();
}

/** يخلّي صور فولدر meals_Images بابلك ويكتب لينكاتها في عمودين جداد */
function setupImages() {
  const it = DriveApp.getFoldersByName(IMAGES_FOLDER);
  if (!it.hasNext()) throw new Error('مش لاقي فولدر ' + IMAGES_FOLDER);
  const folder = it.next();
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  const map = {};
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    map[f.getName()] = f.getId();
  }
  const url = id => id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000' : '';
  const fileId = p => { const n = String(p || '').split('/').pop(); return map[n] || ''; };
  const m = ss_().getSheetByName(SH.MEALS);
  const last = m.getLastRow();
  if (last < 2) return;
  const rows = m.getRange(2, 1, last - 1, 5).getValues();
  const out = rows.map(r => [
    url(fileId(r[1])),
    [r[2], r[3], r[4]].map(fileId).filter(String).map(url).join('\n')
  ]);
  m.getRange(2, META_COL, out.length, 2).setValues(out);
}

/* ============ API ============ */
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'all';
  if (action === 'all') return json_(getAll_());
  return json_({ ok: false, error: 'unknown action' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    switch (b.action) {
      case 'savePlan': savePlan_(b.week, b.rows); break;
      case 'check': saveCheck_(b.scope, b.item, b.checked); break;
      case 'settings': saveSettings_(b.data || {}); break;
      case 'log': appendLog_(b); break;
      case 'seed': seedMeta_(b.meals || []); break;
      default: return json_({ ok: false, error: 'unknown action' });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function getAll_() {
  const m = ss_().getSheetByName(SH.MEALS);
  const last = m.getLastRow();
  const meals = [];
  if (last > 1) {
    const w = Math.max(META_COL + META_HEAD.length - 1, 6);
    m.getRange(2, 1, last - 1, w).getValues().forEach(r => {
      if (!r[0]) return;
      meals.push({
        en: String(r[0]).trim(), link: r[5] && r[5] !== '0' && r[5] !== 0 ? String(r[5]) : '',
        photo: r[6] || '', contents: String(r[7] || '').split('\n').filter(String),
        id: r[8] || '', ar: r[9] || '', seeded: !!r[9]
      });
    });
  }
  const plan = rows_(SH.PLAN).map(r => ({
    week: dstr_(r[0]), date: dstr_(r[1]), slot: r[3], id: r[4], status: r[7] || 'planned', locked: r[8] === true || r[8] === 'TRUE'
  }));
  const shop = {};
  rows_(SH.SHOP).forEach(r => { if (r[2] === true || r[2] === 'TRUE') { (shop[r[0]] = shop[r[0]] || {})[r[1]] = true; } });
  const settings = {};
  rows_(SH.SET).forEach(r => { if (r[0]) settings[r[0]] = r[1]; });
  return { ok: true, meals, plan, shop, settings };
}

function rows_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function dstr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

/** يستبدل صفوف أسبوع كامل */
function savePlan_(week, rows) {
  const sh = sheet_(SH.PLAN, PLAN_HEAD);
  const last = sh.getLastRow();
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, PLAN_HEAD.length).getValues();
    const keep = vals.filter(r => dstr_(r[0]) !== week);
    sh.getRange(2, 1, last - 1, PLAN_HEAD.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, PLAN_HEAD.length).setValues(keep);
  }
  if (!rows || !rows.length) return;
  const now = new Date();
  const out = rows.map(r => ["'" + week, "'" + r.date, r.day, r.slot, r.id, r.ar, r.kcal, r.status || 'planned', !!r.locked, now]);
  sh.getRange(sh.getLastRow() + 1, 1, out.length, PLAN_HEAD.length).setValues(out);
}

function saveCheck_(scope, item, checked) {
  const sh = sheet_(SH.SHOP, ['scope', 'item', 'checked', 'updated_at']);
  const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === scope && vals[i][1] === item) {
      sh.getRange(i + 2, 3, 1, 2).setValues([[!!checked, new Date()]]);
      return;
    }
  }
  sh.appendRow([scope, item, !!checked, new Date()]);
}

function saveSettings_(data) {
  const sh = sheet_(SH.SET, ['key', 'value']);
  const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  Object.keys(data).forEach(k => {
    const i = vals.indexOf(k);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(data[k]);
    else { sh.appendRow([k, data[k]]); vals.push(k); }
  });
}

function appendLog_(b) {
  sheet_(SH.LOG, ['timestamp', 'date', 'slot', 'from', 'to', 'changes'])
    .appendRow([new Date(), "'" + b.date, b.slot, b.from, b.to, b.changes || '']);
}

/** يكتب بيانات الوجبات (عربي/سعرات/مكونات) جنب كل وجبة — بيتبعت من التطبيق أول مرة */
function seedMeta_(meals) {
  const m = ss_().getSheetByName(SH.MEALS);
  const last = m.getLastRow();
  if (last < 2) return;
  const names = m.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim());
  const byEn = {};
  meals.forEach(x => byEn[x.en] = x);
  const out = names.map(n => {
    const x = byEn[n];
    if (!x) return ['', '', '', '', '', '', '', '', ''];
    return [x.id, x.ar, x.slots.join(','), x.group, x.kcal, x.p, x.prep, x.serves, x.ingText];
  });
  m.getRange(2, META_COL + 2, out.length, 9).setValues(out);
}
