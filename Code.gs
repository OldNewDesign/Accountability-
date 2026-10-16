/**
 * 2nd PLT Accountability → Google Sheet
 *
 * SETUP (one time, about 5 minutes)
 *  1. Create a blank Google Sheet. Open Extensions > Apps Script.
 *  2. Delete what's in Code.gs and paste this whole file in.
 *  3. Change ACCESS_CODE below to a code only your leaders know. Save.
 *  4. Deploy > New deployment > gear icon > Web app.
 *       Execute as: Me      Who has access: Anyone
 *     Click Deploy, allow access, and copy the Web app URL.
 *  5. In the app: Roster tab > Google Sheet. Paste the URL and the access code, tap Connect.
 *
 * If you change this script later: Deploy > Manage deployments > pencil icon >
 * Version: New version > Deploy. The URL stays the same.
 */
const ACCESS_CODE = 'CHANGE-ME';
const TIMEZONE = 'Asia/Seoul';   // times on the report tabs
const PULL_DAYS = 45;            // days sent back to each phone when it syncs

const LABEL = {present:'Present', appt:'Appt', pass:'Pass', leave:'Leave', detail:'Detail', late:'Late', other:'Other'};
const ORDER = ['appt','pass','leave','detail','late','other'];
const COLOR = {present:'#2d7a3c', appt:'#6a4cae', pass:'#2c64a1', leave:'#2c64a1', detail:'#566270', late:'#a86b12', other:'#8a5527', none:'#b3261e'};
const MON = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const DOW = ['SUN','MON','TUE','WED','THU','FRI','SAT'];

function doGet(e) { return handle_((e && e.parameter) || {}, null); }

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ok:false, error:'bad_request'}); }
  return handle_((e && e.parameter) || {}, body);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(params, body) {
  if (ACCESS_CODE === 'CHANGE-ME') return out_({ok:false, error:'not_configured'});
  const code = String((body && body.code) || params.code || '');
  if (code !== ACCESS_CODE) return out_({ok:false, error:'bad_code'});

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const st = store_();
    if (body && body.action === 'push') {
      if (body.roster && body.roster.sections) {
        const cur = st.get('config/roster');
        if (!cur || (body.roster.updatedAt || '') >= (cur.updatedAt || '')) st.put('config/roster', body.roster);
      }
      const changed = [];
      (body.days || []).forEach(function (d) {
        if (!d || !/^\d{4}-\d{2}-\d{2}$/.test(d.date || '')) return;
        const m = mergeDay(st.get('days/' + d.date), d);
        st.put('days/' + d.date, m);
        changed.push(m);
      });
      st.flush();
      changed.sort(function (a, b) { return a.date < b.date ? -1 : 1; }).forEach(writeDayTab_);
      if (changed.length) writeSummary_(st);
      return out_({ok:true, roster:st.get('config/roster'), days:changed});
    }
    // pull
    const cutoff = Utilities.formatDate(new Date(Date.now() - PULL_DAYS * 864e5), TIMEZONE, 'yyyy-MM-dd');
    const days = st.keys('days/').map(st.get).filter(function (d) { return d && d.date >= cutoff; });
    return out_({ok:true, roster:st.get('config/roster'), days:days});
  } finally {
    lock.releaseLock();
  }
}

/* Two phones editing the same day: keep each Soldier's most recent mark. Same function runs in the app. */
function mergeDay(a, b) {
  a = a || {}; b = b || {};
  const newer = (b.updatedAt || '') > (a.updatedAt || '') ? b : a;
  const out = JSON.parse(JSON.stringify(newer));
  const fa = a.formations || {}, fb = b.formations || {};
  out.formations = {};
  Object.keys(fa).concat(Object.keys(fb)).forEach(function (fid) {
    if (out.formations[fid]) return;
    const x = fa[fid] || {}, y = fb[fid] || {};
    const starts = [x.startedAt, y.startedAt].filter(Boolean).sort();
    const f = {label:(newer === b ? (y.label || x.label) : (x.label || y.label)) || '', startedAt:starts[0] || '', entries:{}};
    const ex = x.entries || {}, ey = y.entries || {};
    Object.keys(ex).concat(Object.keys(ey)).forEach(function (sid) {
      const p = ex[sid], q = ey[sid];
      f.entries[sid] = !p ? q : !q ? p : ((q.at || '') > (p.at || '') ? q : p);
    });
    out.formations[fid] = f;
  });
  out.updatedAt = [a.updatedAt, b.updatedAt].filter(Boolean).sort().pop() || '';
  return out;
}

/* Raw records live on a hidden _data tab: one row per day. */
function store_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('_data');
  if (!sh) {
    sh = ss.insertSheet('_data');
    sh.getRange(1, 1, 1, 3).setValues([['key', 'json', 'updated']]);
    sh.hideSheet();
  }
  const n = sh.getLastRow();
  const vals = n > 1 ? sh.getRange(2, 1, n - 1, 2).getValues() : [];
  const rows = {}, cache = {}, dirty = {};
  vals.forEach(function (r, i) { if (r[0]) { rows[r[0]] = i + 2; cache[r[0]] = r[1]; } });
  let next = Math.max(n, 1) + 1;
  return {
    get: function (k) { const v = cache[k]; if (!v) return null; try { return JSON.parse(v); } catch (e) { return null; } },
    put: function (k, o) { cache[k] = JSON.stringify(o); dirty[k] = true; },
    keys: function (prefix) { return Object.keys(cache).filter(function (k) { return k.indexOf(prefix) === 0; }); },
    flush: function () {
      Object.keys(dirty).forEach(function (k) {
        let r = rows[k];
        if (!r) { r = next++; rows[k] = r; }
        sh.getRange(r, 1, 1, 3).setValues([[k, cache[k], new Date()]]);
      });
    }
  };
}

/* ---------- report helpers ---------- */
function tabName_(date) { const p = date.split('-'); return p[2] + ' ' + MON[+p[1] - 1] + ' ' + p[0]; }
function dow_(date) { const p = date.split('-').map(Number); return DOW[new Date(p[0], p[1] - 1, p[2]).getDay()]; }
function hhmm_(iso) { return iso ? Utilities.formatDate(new Date(iso), TIMEZONE, 'HHmm') : ''; }
function formations_(d) {
  const f = d.formations || {};
  return Object.keys(f).map(function (id) { const o = f[id]; return {id:id, label:o.label || '', startedAt:o.startedAt || '', entries:o.entries || {}}; })
    .sort(function (a, b) { return a.startedAt < b.startedAt ? -1 : a.startedAt > b.startedAt ? 1 : 0; });
}
function statusText_(e) {
  if (!e || !e.s) return 'Not accounted';
  return e.s === 'other' ? (String(e.other || '').trim() || 'Other') : LABEL[e.s];
}
function tally_(d, ents) {
  const c = {assigned:0, present:0, open:0, st:{}, by:{}};
  ((d.roster || {}).sections || []).forEach(function (sec) {
    sec.soldiers.forEach(function (p) {
      c.assigned++;
      const e = ents[p.id] || {}, s = e.s || '';
      if (!s) { c.open++; return; }
      if (s === 'present') { c.present++; return; }
      c.st[s] = (c.st[s] || 0) + 1;
      const name = statusText_(e), k = s === 'other' ? 'other:' + name.toLowerCase() : s;
      if (!c.by[k]) c.by[k] = {name:name, st:s, n:0};
      c.by[k].n++;
    });
  });
  c.list = Object.keys(c.by).map(function (k) { return c.by[k]; })
    .sort(function (a, b) { return (ORDER.indexOf(a.st) - ORDER.indexOf(b.st)) || (a.name < b.name ? -1 : 1); });
  return c;
}

function summarySheet_(ss) {
  let sh = ss.getSheetByName('Summary');
  if (sh) return sh;
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0) return s1.setName('Summary');
  return ss.insertSheet('Summary', 0);
}

/* One tab per day, newest first, rewritten as the day goes on. The final formation is on top. */
function writeDayTab_(d) {
  const ss = SpreadsheetApp.getActive();
  summarySheet_(ss);
  const name = tabName_(d.date);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, 1);

  const rows = [], bg = [], fc = [], fw = [];
  function add(r, opt) {
    opt = opt || {};
    while (r.length < 4) r.push('');
    rows.push(r.map(function (v) { return v == null ? '' : String(v); }));
    const b = opt.bg || '#ffffff';
    bg.push([b, b, b, b]);
    const f = opt.fc || '#1c2018';
    fc.push([f, opt.fcB || f, opt.fcC || f, f]);
    const w = opt.bold ? 'bold' : 'normal';
    fw.push([w, w, w, w]);
  }

  const fs = formations_(d);
  const unit = d.unit || (d.roster && d.roster.unit) || 'PLT';
  add([unit + ' — End of Day Report'], {bold:true});
  add(['DATE:', dow_(d.date) + ' ' + tabName_(d.date)], {bold:true});
  add(['Updated:', d.updatedAt ? Utilities.formatDate(new Date(d.updatedAt), TIMEZONE, 'dd MMM yyyy HHmm').toUpperCase() : '']);
  add([]);

  if (!fs.length) {
    add(['No formation recorded.']);
  } else {
    const fin = fs[fs.length - 1];
    const c = tally_(d, fin.entries);
    add(['NUMBERS', fin.label + ' ' + hhmm_(fin.startedAt) + ' (final)'], {bold:true});
    add([c.assigned, 'Assigned']);
    add([c.present, 'Present'], {fcB:COLOR.present});
    c.list.forEach(function (b) { add([b.n, b.name], {fcB:COLOR[b.st]}); });
    if (c.open) add([c.open, 'Not accounted'], {fcB:COLOR.none});
    add([]);

    fs.slice().reverse().forEach(function (f, i) {
      const ents = f.entries;
      add([f.label + ' · ' + hhmm_(f.startedAt) + (i === 0 ? ' (final)' : '')], {bold:true, bg:'#d9dccf'});
      add(['Rank', 'Last Name', 'Status', 'Remarks'], {bold:true, bg:'#465a2a', fc:'#ffffff'});
      ((d.roster || {}).sections || []).forEach(function (sec) {
        const pres = sec.soldiers.filter(function (p) { return (ents[p.id] || {}).s === 'present'; }).length;
        add([sec.name, '', pres + ' / ' + sec.soldiers.length + ' present'], {bold:true, bg:'#eceee6'});
        sec.soldiers.forEach(function (p) {
          const e = ents[p.id] || {};
          add([p.rank, p.name, statusText_(e), e.rmk || ''], {fcC:e.s ? COLOR[e.s] : COLOR.none});
        });
      });
      add([]);
    });
  }

  sh.clear();
  const rg = sh.getRange(1, 1, rows.length, 4);
  rg.setNumberFormat('@');
  rg.setValues(rows);
  rg.setBackgrounds(bg);
  rg.setFontColors(fc);
  rg.setFontWeights(fw);
  sh.getRange(1, 1).setFontSize(14);
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(3, 190);
  sh.setColumnWidth(4, 320);
}

/* Summary: one row per day, newest first, numbers from the day's final formation. */
function writeSummary_(st) {
  const ss = SpreadsheetApp.getActive();
  const sh = summarySheet_(ss);
  const ids = {};
  ss.getSheets().forEach(function (s) { ids[s.getName()] = s.getSheetId(); });
  const days = st.keys('days/').map(st.get).filter(Boolean).sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  const head = ['Date', 'Day', 'Final formation', 'Assigned', 'Present', 'Appt', 'Pass', 'Leave', 'Detail', 'Late', 'Other', 'Not accounted', 'Report'];
  const rows = days.map(function (d) {
    const fs = formations_(d), f = fs[fs.length - 1];
    const c = tally_(d, f ? f.entries : {});
    const name = tabName_(d.date);
    return [name, dow_(d.date), f ? (f.label + ' ' + hhmm_(f.startedAt)) : '', c.assigned, c.present,
      c.st.appt || 0, c.st.pass || 0, c.st.leave || 0, c.st.detail || 0, c.st.late || 0, c.st.other || 0, c.open,
      ids[name] != null ? '=HYPERLINK("#gid=' + ids[name] + '","Open")' : ''];
  });
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#465a2a').setFontColor('#ffffff');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  }
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(3, 190);
}
