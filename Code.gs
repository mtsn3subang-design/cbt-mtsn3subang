/* CBT MTsN 3 Subang – Backend Apps Script (multi mapel & multi guru) — VERSI OPTIMASI BEBAN TINGGI
   Database: Google Sheets di bawah. Sheet dibuat otomatis saat pertama dipakai.
   Perubahan utama: cache (CacheService) untuk tabel yang jarang berubah, tanpa kunci global untuk simpan jawaban,
   pencarian hasil per baris (tidak membaca seluruh sheet), dan kunci hanya sebentar saat mulai tes. */
const SS_ID = '1zWHSga-6ossQvT85ZV_G4jH5SJsxTgCUAH21-XSh-YM';
const TABLES = ['mapel', 'guru', 'topik', 'soal', 'grp', 'pes', 'tes', 'hasil'];
const GURU_TABLES = ['topik', 'soal', 'tes', 'hasil'];
const TZ = 'Asia/Jakarta';
const CACHE_TTL = 600;            // detik; cache otomatis dibuang saat data diubah lewat aplikasi
const NOCACHE = { hasil: 1 };      // tabel yang selalu dibaca langsung
const LOCKED = ['sync', 'gprof', 'imgup', 'imgclean', 'export', 'sinkron']; // aksi yang memakai kunci global

/* ---------- util ---------- */
const ok = o => out(Object.assign({ ok: true }, o || {}));
const err = m => out({ ok: false, err: m });
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const today = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const uid = () => Utilities.getUuid().replace(/-/g, '').slice(0, 7);
const shuf = a => { a = a.slice(); for (let i = a.length; i > 1;) { const j = Math.floor(Math.random() * i--); [a[i], a[j]] = [a[j], a[i]]; } return a; };
const summ = o => String(o.nama || o.q || o.u || (o.tes ? 'tes ' + o.tes + ' | nilai ' + o.nilai : '') || '').slice(0, 80);

let _ss = null; const _sh = {};
function sh(n) {
  if (_sh[n]) return _sh[n];
  if (!_ss) _ss = SpreadsheetApp.openById(SS_ID);
  let x = _ss.getSheetByName(n);
  if (!x) { x = _ss.insertSheet(n); x.getRange('A:C').setNumberFormat('@'); x.getRange(1, 1, 1, 3).setValues([['id', 'data (JSON)', 'ringkas']]); x.setFrozenRows(1); }
  return (_sh[n] = x);
}
function ensureRows(s, need) { // pastikan baris sheet cukup (kolom A:C tetap berformat teks)
  const mx = s.getMaxRows();
  if (need <= mx) return;
  const add = Math.max(500, need - mx);
  s.insertRowsAfter(mx, add);
  s.getRange(mx + 1, 1, add, 3).setNumberFormat('@');
}

/* ---------- cache berversi (chunk 30 rb karakter, aman untuk batas 100 KB/kunci) ---------- */
const newVer = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
function cver(c, name) { let v = c.get('v:' + name); if (!v) { v = newVer(); c.put('v:' + name, v, 21600); } return v; }
function cbump(name) { try { CacheService.getScriptCache().put('v:' + name, newVer(), 21600); } catch (e) {} }
function cget(name, loader, ttl, ver) {
  const c = CacheService.getScriptCache();
  const v = ver || cver(c, name), base = 'd:' + name + ':' + v;
  try {
    const head = c.get(base + ':n');
    if (head !== null) {
      const n = +head, keys = [];
      for (let i = 0; i < n; i++) keys.push(base + ':' + i);
      let s = '', all = n > 0;
      for (let i = 0; i < keys.length && all; i += 40) {
        const sub = keys.slice(i, i + 40), got = c.getAll(sub);
        for (const k of sub) { const p = got[k]; if (p === undefined || p === null) { all = false; break; } s += p; }
      }
      if (all) return JSON.parse(s);
    }
  } catch (e) {}
  const data = loader();
  if (data === undefined) return data;
  try {
    const s = JSON.stringify(data), CH = 30000, put = {}, ks = [];
    let n = 0;
    for (let i = 0; i < s.length; i += CH, n++) { put[base + ':' + n] = s.slice(i, i + CH); ks.push(base + ':' + n); }
    if (n <= 120) {
      put[base + ':n'] = String(n); ks.push(base + ':n'); // header terakhir: cache baru dianggap utuh jika header ada
      for (let i = 0; i < ks.length; i += 40) { const sub = {}; ks.slice(i, i + 40).forEach(k => sub[k] = put[k]); c.putAll(sub, ttl || CACHE_TTL); }
    }
  } catch (e) {}
  return data;
}
function resetCache() { // jalankan manual dari editor Apps Script jika Anda mengedit sheet langsung
  TABLES.forEach(n => cbump('t:' + n)); cbump('kv:cfg'); cbump('kv:token');
}
const imgIds = s => [s.img].concat(s.oi || []).filter(Boolean);
function removeImgs(ids) {
  if (!ids.length) return;
  const s = sh('gambar'), last = s.getLastRow(); if (last < 2) return;
  const v = s.getRange(2, 1, last - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (ids.indexOf(String(v[i][0])) >= 0) s.deleteRow(i + 2);
}
function writeSheet(name, rows) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let s = ss.getSheetByName(name); if (!s) s = ss.insertSheet(name);
  s.clear();
  const w = Math.max.apply(null, rows.map(r => r.length));
  const v = rows.map(r => { const a = r.slice(); while (a.length < w) a.push(''); return a; });
  const rg = s.getRange(1, 1, v.length, w); rg.setNumberFormat('@'); rg.setValues(v);
  s.getRange(1, 1, 1, 1).setFontWeight('bold');
}
function extractId(s) { const m = String(s || '').match(/[-\w]{25,}/); return m ? m[0] : ''; }
function setByLabel(s, label, val) {
  const v = s.getRange(1, 1, 40, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === label) { s.getRange(i + 1, 2).setValue(val); return true; }
  return false;
}
function readRaw(n) {
  const s = sh(n), last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 2, last - 1, 1).getValues().map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } }).filter(Boolean);
}
function readT(n) { return NOCACHE[n] ? readRaw(n) : cget('t:' + n, () => readRaw(n)); }
const soalMap = () => { const m = {}; readT('soal').forEach(s => m[s.id] = s); return m; };
const tesById = id => readT('tes').find(x => x.id === id) || null;
function upsert(n, list) {
  if (!list.length) return;
  const s = sh(n), last = s.getLastRow(), idx = {}, add = [];
  if (last > 1) s.getRange(2, 1, last - 1, 1).getValues().forEach((r, i) => idx[r[0]] = i + 2);
  list.forEach(o => {
    const row = [o.id, JSON.stringify(o), summ(o)];
    if (idx[o.id]) s.getRange(idx[o.id], 1, 1, 3).setValues([row]); else add.push(row);
  });
  if (add.length) { ensureRows(s, s.getLastRow() + add.length); s.getRange(s.getLastRow() + 1, 1, add.length, 3).setValues(add); }
  cbump('t:' + n);
}
function remove(n, ids) {
  const s = sh(n), last = s.getLastRow();
  if (last < 2 || !ids.length) return;
  const v = s.getRange(2, 1, last - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (ids.indexOf(v[i][0]) >= 0) s.deleteRow(i + 2);
  cbump('t:' + n);
}
function getOne(n, id) {
  const f = sh(n).getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!f) return null;
  return { row: f.getRow(), obj: JSON.parse(sh(n).getRange(f.getRow(), 2).getValue()) };
}
function putOne(n, row, o) { sh(n).getRange(row, 1, 1, 3).setValues([[o.id, JSON.stringify(o), summ(o)]]); cbump('t:' + n); }

/* ---------- hasil: akses per baris tanpa membaca seluruh sheet ---------- */
function hasilRows(pid) { // semua hasil milik satu peserta
  const s = sh('hasil'), last = s.getLastRow(); if (last < 2) return [];
  const fs = s.getRange(2, 2, last - 1, 1).createTextFinder('"p":"' + pid + '"').findAll(), out = [];
  fs.forEach(f => { try { const o = JSON.parse(f.getValue()); if (o && o.p === pid) out.push({ row: f.getRow(), obj: o }); } catch (e) {} });
  return out;
}
// Ubah satu hasil tanpa kunci global: baca, ubah, cek ulang baris belum berubah, baru tulis (ulang jika bentrok).
function updHasil(id, fn) {
  if (!id) return { miss: true };
  const s = sh('hasil'), c = CacheService.getScriptCache();
  let row = +c.get('hr:' + id) || 0;
  for (let t = 0; t < 4; t++) {
    let v = null;
    if (row > 1) { try { v = s.getRange(row, 1, 1, 2).getValues()[0]; } catch (e) { v = null; } }
    if (!v || String(v[0]) !== id) {
      const f = s.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
      if (!f) return { miss: true };
      row = f.getRow(); c.put('hr:' + id, String(row), 21600);
      v = s.getRange(row, 1, 1, 2).getValues()[0];
      if (String(v[0]) !== id) continue;
    }
    let h; try { h = JSON.parse(v[1]); } catch (e) { return { miss: true }; }
    if (fn(h) === false) return { h: h, skipped: true };
    const cur = s.getRange(row, 1, 1, 2).getValues()[0];
    if (String(cur[0]) !== id || cur[1] !== v[1]) continue;
    s.getRange(row, 2, 1, 2).setValues([[JSON.stringify(h), summ(h)]]);
    return { h: h, ok: true };
  }
  return { busy: true };
}
function getKV(k) {
  return cget('kv:' + k, () => { const f = sh('cfg').getRange('A:A').createTextFinder(k).matchEntireCell(true).findNext(); return f ? JSON.parse(sh('cfg').getRange(f.getRow(), 2).getValue()) : null; });
}
function setKV(k, v) {
  const s = sh('cfg'), f = s.getRange('A:A').createTextFinder(k).matchEntireCell(true).findNext();
  if (f) s.getRange(f.getRow(), 2).setValue(JSON.stringify(v)); else s.appendRow([k, JSON.stringify(v)]);
  cbump('kv:' + k);
}

/* ---------- auth ---------- */
function who(a) {
  if (!a) return null;
  const c = getKV('cfg') || {};
  if (a.u === (c.au || 'admin') && a.p === (c.ap || 'admin123')) return { role: 'admin' };
  const g = readT('guru').find(x => x.u === a.u && x.p === a.p);
  return g ? { role: 'guru', g: g } : null;
}
let _sec = null;
function secret() {
  if (_sec) return _sec;
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SEC');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SEC', s); }
  return (_sec = s);
}
const sign = pes => pes.id + '.' + Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(pes.id + '|' + pes.p, secret())).replace(/=+$/, '');
function pesFrom(tk) {
  tk = String(tk || ''); if (!tk || tk.length > 200) return null;
  const c = CacheService.getScriptCache(), ck = 'pk:' + tk;
  try { const h = c.get(ck); if (h) return JSON.parse(h); } catch (e) {}
  const id = tk.split('.')[0], o = readT('pes').find(x => x.id === id);
  if (o && sign(o) === tk) { try { c.put(ck, JSON.stringify({ id: o.id, g: o.g }), 300); } catch (e) {} return o; }
  return null;
}

/* ---------- data ---------- */
function tokMap() { const t = getKV('token') || {}; return t.c !== undefined ? {} : t; }
function cfgPub() { const c = getKV('cfg') || {}; return { nama: c.nama || 'CBT MTsN 3 Subang', ket: c.ket || 'Ujian Online Berbasis Komputer', kepala: c.kepala || '', nipk: c.nipk || '', tempat: c.tempat || 'Subang' }; }
function buildS(w) {
  const c = getKV('cfg') || {}, cfg = Object.assign(cfgPub(), { au: c.au || 'admin', tpl: c.tpl || '' });
  const S = { cfg, mapel: readT('mapel'), guru: readT('guru'), topik: readT('topik'), soal: readT('soal'), grp: readT('grp'), pes: readT('pes'), tes: readT('tes'), hasil: readT('hasil'), token: tokMap(), role: w.role };
  if (w.role === 'admin') { if (!c.ap) S.warn = 'Akun admin masih default (admin / admin123). Segera ganti di Pengaturan & Backup.'; return S; }
  const m = w.g.m || [];
  S.role = 'guru'; delete S.cfg.au; delete S.cfg.tpl;
  S.guru = [w.g];
  S.token = Object.fromEntries(Object.entries(S.token).filter(e => m.indexOf(e[0]) >= 0));
  S.mapel = S.mapel.filter(x => m.indexOf(x.id) >= 0);
  S.topik = S.topik.filter(t => m.indexOf(t.m) >= 0);
  const tid = {}; S.topik.forEach(t => tid[t.id] = 1);
  S.soal = S.soal.filter(s => tid[s.topik]);
  S.tes = S.tes.filter(t => m.indexOf(t.m) >= 0);
  const eid = {}; S.tes.forEach(t => eid[t.id] = 1);
  S.hasil = S.hasil.filter(h => eid[h.tes]);
  S.pes = S.pes.map(p => ({ id: p.id, nama: p.nama, g: p.g }));
  return S;
}
function score(h, T, sm) {
  sm = sm || soalMap();
  let p = 0;
  h.ids.forEach(i => {
    const s = sm[i]; if (!s) return;
    const a = h.a[i];
    p += (a === undefined || a === null) ? (+T.pk || 0) : (a === s.k ? (+T.pb || 0) : (+T.ps || 0));
  });
  return Math.round(p / (h.ids.length * (+T.pb || 1)) * 1000) / 10; // nilai skala 0-100
}
const endMs = (h, T) => { let e = h.mulai + (T.durasi + (h.x || 0)) * 60000; if (T.selesai) e = Math.min(e, new Date(T.selesai).getTime() + (h.x || 0) * 60000); return e; };

/* ---------- sync admin/guru ---------- */
function sync(a, ch) {
  const w = who(a); if (!w) return err('Login tidak valid.');
  const gu = w.role === 'guru', myM = gu ? (w.g.m || []) : null;
  if (!gu && ch.cfg) {
    const c = getKV('cfg') || {}, n = ch.cfg;
    setKV('cfg', { nama: n.nama, ket: n.ket, kepala: n.kepala || '', nipk: n.nipk || '', tempat: n.tempat !== undefined ? n.tempat : (c.tempat || ''), au: n.au || c.au || 'admin', ap: n.ap || c.ap || '', tpl: n.tpl !== undefined ? n.tpl : (c.tpl || '') });
  }
  if (ch.token) { const tm = tokMap(); Object.keys(ch.token).forEach(k => { if (!gu || myM.indexOf(k) >= 0) tm[k] = ch.token[k]; }); setKV('token', tm); }
  for (const t of TABLES) {
    if (!ch[t] || (gu && GURU_TABLES.indexOf(t) < 0)) continue;
    const up = ch[t].up || [], rm = ch[t].rm || [];
    const ex = {}; if (t !== 'hasil') readT(t).forEach(o => ex[o.id] = o);
    let allowed = () => true;
    if (gu) {
      const topikOk = {}, tesOk = {};
      readT('topik').forEach(x => { if (myM.indexOf(x.m) >= 0) topikOk[x.id] = 1; });
      readT('tes').forEach(x => { if (myM.indexOf(x.m) >= 0) tesOk[x.id] = 1; });
      allowed = o => t === 'topik' ? myM.indexOf(o.m) >= 0 : t === 'soal' ? !!topikOk[o.topik] : t === 'tes' ? myM.indexOf(o.m) >= 0 : !!tesOk[o.tes];
    }
    let list = up.filter(o => allowed(o) && (!ex[o.id] || allowed(ex[o.id])));
    if (t === 'hasil') { // diubah per baris agar tidak menimpa jawaban siswa yang sedang tersimpan
      const tesAll = {}; readT('tes').forEach(x => tesAll[x.id] = x);
      const sm = soalMap();
      list.forEach(r => updHasil(String(r.id || ''), h => {
        if (!allowed(h)) return false;
        h.fin = !!(h.fin || r.fin); h.x = +r.x || 0;
        if (h.fin && tesAll[h.tes]) h.nilai = score(h, tesAll[h.tes], sm);
      }));
      rm.forEach(id => { const g = getOne('hasil', String(id)); if (g && allowed(g.obj)) sh('hasil').deleteRow(g.row); });
      continue;
    }
    const rmOk = rm.filter(id => ex[id] && allowed(ex[id]));
    if (t === 'soal') {
      const gone = [];
      list.forEach(o => { const e = ex[o.id]; if (e) { const keep = imgIds(o); imgIds(e).forEach(i => { if (keep.indexOf(i) < 0) gone.push(i); }); } });
      rmOk.forEach(id => imgIds(ex[id]).forEach(i => gone.push(i)));
      removeImgs(gone);
    }
    upsert(t, list);
    remove(t, rmOk);
  }
  return ok();
}

/* ---------- entry ---------- */
function doGet() { return ContentService.createTextOutput('CBT API OK'); }
function doPost(e) {
  let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return err('Permintaan tidak valid.'); }
  const needLock = LOCKED.indexOf(b.action) >= 0, lock = LockService.getScriptLock();
  try {
    if (needLock && !lock.tryLock(25000)) return err('Server sibuk, coba lagi.');
    return route(b);
  } catch (x) { return err('Kesalahan server: ' + x.message); }
  finally { if (needLock) try { lock.releaseLock(); } catch (z) {} }
}
function route(b) {
  switch (b.action) {
    case 'info': return ok({ cfg: cfgPub() });
    case 'alogin': { const w = who({ u: b.u, p: b.p }); return w ? ok({ S: buildS(w) }) : err('Username atau password salah.'); }
    case 'sync': return sync(b.auth, b.ch || {});
    case 'plogin': {
      const pes = readT('pes').find(x => x.u === b.u && x.p === b.p);
      if (!pes) return err('Username atau password salah.');
      return ok({ now: Date.now(), tk: sign(pes), pes: { id: pes.id, u: pes.u, nama: pes.nama, g: pes.g }, cfg: cfgPub(), grp: readT('grp'), mapel: readT('mapel'),
        tes: readT('tes').filter(t => (t.g || []).indexOf(pes.g) >= 0), hasil: hasilRows(pes.id).map(x => x.obj) });
    }
    case 'pstart': {
      const pes = pesFrom(b.tk); if (!pes) return err('Sesi tidak valid, silakan login ulang.');
      const T = tesById(b.tes); if (!T || (T.g || []).indexOf(pes.g) < 0) return err('Tes tidak tersedia.');
      const now = Date.now(), sm = soalMap();
      const pack = h => ok({ now: Date.now(), h, soal: h.ids.filter(i => sm[i]).map(i => ({ id: i, topik: sm[i].topik, q: sm[i].q, o: sm[i].o, img: sm[i].img || '', oi: sm[i].oi || [] })) });
      let ex = hasilRows(pes.id).find(x => x.obj.tes === T.id);
      if (ex) { if (ex.obj.fin) return err('Tes sudah selesai.'); return pack(ex.obj); }
      if (T.mulai && now < new Date(T.mulai).getTime()) return err('Tes belum dibuka.');
      if (T.selesai && now > new Date(T.selesai).getTime()) return err('Tes sudah ditutup.');
      if (T.token) { const tk = tokMap()[T.m] || {}; if (!b.token || String(b.token).trim().toUpperCase() !== tk.c || tk.d !== today()) return err('Token salah atau kedaluwarsa.'); }
      if (!T.ids || !T.ids.length) return err('Soal belum tersedia.');
      const ids = (T.as ? shuf(T.ids) : T.ids.slice()).filter(i => sm[i]), perm = {};
      ids.forEach(i => { const n = sm[i].o.map((_, k) => k); perm[i] = T.aj ? shuf(n) : n; });
      const h = { id: uid(), tes: T.id, p: pes.id, mulai: now, x: 0, fin: false, nilai: 0, a: {}, ids, perm };
      // kunci HANYA untuk memesan nomor baris (±0,1 dtk); penulisan data dilakukan di luar kunci
      const cch = CacheService.getScriptCache(), mk = 'st:' + pes.id + ':' + T.id;
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(25000)) return err('Server sibuk, coba lagi.');
      let row;
      try {
        const m = cch.get(mk); // penanda: siswa ini sudah/sedang dibuatkan hasil (permintaan ganda/ulang)
        if (m) {
          const pr0 = m.split(':'), v = sh('hasil').getRange(+pr0[1], 1, 1, 2).getValues()[0];
          if (String(v[0]) !== pr0[0]) return err('Server sibuk, coba lagi.'); // masih ditulis oleh permintaan pertama
          let eo = null; try { eo = JSON.parse(v[1]); } catch (e3) {}
          if (eo) { if (eo.fin) return err('Tes sudah selesai.'); return pack(eo); }
        }
        const s = sh('hasil'), pr = PropertiesService.getScriptProperties();
        const last = s.getLastRow(), pend = +pr.getProperty('HROW') || 0;
        row = ((pend > last && pend - last <= 60) ? pend : last) + 1; // 60 = batas wajar baris yang sedang ditulis
        ensureRows(s, row);
        pr.setProperty('HROW', String(row));
        cch.put(mk, h.id + ':' + row, 30); // sementara 30 dtk; jika penulisan gagal, otomatis kedaluwarsa
      } finally { lock.releaseLock(); }
      sh('hasil').getRange(row, 1, 1, 3).setValues([[h.id, JSON.stringify(h), summ(h)]]);
      try { cch.put('hr:' + h.id, String(row), 21600); cch.put(mk, h.id + ':' + row, 21600); } catch (e2) {}
      return pack(h);
    }
    case 'imgup': {
      const w = who(b.auth); if (!w) return err('Login tidak valid.');
      const d = String(b.d || '');
      if (d.indexOf('data:image/') !== 0 || d.length > 300000) return err('Gambar tidak valid atau terlalu besar.');
      const id = 'i' + uid() + uid(), s = sh('gambar'), rows = [];
      for (let i = 0, n = 0; i < d.length; i += 45000, n++) rows.push([id, String(n), d.slice(i, i + 45000)]);
      ensureRows(s, s.getLastRow() + rows.length);
      s.getRange(s.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
      return ok({ id });
    }
    case 'img': {
      if (b.auth ? !who(b.auth) : !pesFrom(b.tk)) return err('Tidak diizinkan.');
      const id = String(b.id || '');
      const d = cget('img:' + id, () => {
        const s = sh('gambar'), f = s.getRange('A:A').createTextFinder(id).matchEntireCell(true).findAll();
        if (!f.length) return undefined;
        const rows = f.map(x => x.getRow()), r0 = Math.min.apply(null, rows), r1 = Math.max.apply(null, rows);
        const v = s.getRange(r0, 1, r1 - r0 + 1, 3).getValues();
        return v.filter(r => String(r[0]) === id).map(r => ({ n: parseInt(r[1], 10), t: String(r[2]) })).sort((a, c) => a.n - c.n).map(p => p.t).join('');
      }, 21600, '1');
      if (d === undefined) return err('Gambar tidak ditemukan.');
      return ok({ d: d });
    }
    case 'imgclean': {
      const w = who(b.auth); if (!w || w.role !== 'admin') return err('Hanya admin.');
      const keep = {}; readT('soal').forEach(s => imgIds(s).forEach(i => keep[i] = 1));
      const s = sh('gambar'), last = s.getLastRow(), seen = {}; let n = 0;
      if (last > 1) { const v = s.getRange(2, 1, last - 1, 1).getValues(); for (let i = v.length - 1; i >= 0; i--) { const id = String(v[i][0]); if (!keep[id]) { if (!seen[id]) { seen[id] = 1; n++; } s.deleteRow(i + 2); } } }
      return ok({ n });
    }
    case 'export': {
      const w = who(b.auth); if (!w) return err('Login tidak valid.');
      const T = (getOne('tes', b.tes) || {}).obj; if (!T) return err('Tes tidak ditemukan.');
      if (w.role === 'guru' && (w.g.m || []).indexOf(T.m) < 0) return err('Tes ini bukan mapel Anda.');
      const d = b.d || {}, mp = ((getOne('mapel', T.m) || {}).obj || {}).nama || '-';
      const nm = s => String(s).replace(/[\[\]\*\?:\/\\]/g, ' ').slice(0, 30), base = nm(mp) + ' - ' + nm(T.nama) + (d.kelas ? ' - ' + nm(d.kelas) : '');
      const meta = [['Mapel', mp], ['Tes', T.nama], ['Kelas', d.kelas || '-'], ['KKM', d.kkm], ['Peserta selesai', d.n], ['Rata-rata', d.avg], ['Diekspor', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')], ['']];
      const sA = 'Analisis ' + base, sR = 'Remedial ' + base;
      writeSheet(sA, [['ANALISIS BUTIR SOAL']].concat(meta, [['No', 'Soal', '% Benar', 'Kesukaran', 'Daya Pembeda', 'Ket. Daya Pembeda']], d.items || []));
      writeSheet(sR, [['PROGRAM REMEDIAL DAN PENGAYAAN']].concat(meta,
        [['REMEDIAL (nilai di bawah KKM)'], ['Nama', 'Group', 'Nilai', 'No. soal salah']], d.rem || [],
        [[''], ['PENGAYAAN (nilai mencapai KKM)'], ['Nama', 'Group', 'Nilai', 'No. soal salah']], d.pen || [],
        [[''], ['BUTIR SOAL PERLU PENGUATAN (< 60% benar)'], ['No', 'Soal', '% Benar']], d.lemah || [],
        [[''], ['RENCANA KEGIATAN REMEDIAL']], (d.rr || []).map((t, i) => [i + 1, t]),
        [[''], ['RENCANA KEGIATAN PENGAYAAN']], (d.rp || []).map((t, i) => [i + 1, t])));
      return ok({ sheets: [sA, sR] });
    }
    case 'sinkron': {
      const w = who(b.auth); if (!w) return err('Login tidak valid.');
      const T = (getOne('tes', b.tes) || {}).obj; if (!T) return err('Tes tidak ditemukan.');
      if (w.role === 'guru' && (w.g.m || []).indexOf(T.m) < 0) return err('Tes ini bukan mapel Anda.');
      const c = getKV('cfg') || {}, tid = extractId(c.tpl);
      if (!tid) return err('Link template analisis belum diisi. Admin: Pengaturan & Backup → Ubah identitas.');
      let ss; try { ss = SpreadsheetApp.openById(tid); } catch (x) { return err('Template tidak bisa dibuka. Periksa link dan pastikan akun pemilik skrip punya akses edit.'); }
      const sI = ss.getSheetByName('IDENTITAS'), sK = ss.getSheetByName('KUNCI'), sJ = ss.getSheetByName('JAWABAN');
      if (!sI || !sK || !sJ) return err('Sheet IDENTITAS/KUNCI/JAWABAN tidak ditemukan di template.');
      const grp = readT('grp').find(x => x.id === b.g); if (!grp) return err('Pilih kelas terlebih dahulu.');
      const sm = {}; readT('soal').forEach(s => sm[s.id] = s);
      const tm = {}; readT('topik').forEach(x => tm[x.id] = x);
      const ids = (T.ids || []).filter(i => sm[i]).slice(0, 40);
      if (!ids.length) return err('Tes ini belum memiliki soal.');
      const L = 'ABCDE', warn = [];
      if ((T.ids || []).length > 40) warn.push('Soal lebih dari 40: hanya 40 soal pertama yang dikirim.');
      if (ids.some(i => sm[i].o.length > 4)) warn.push('Ada soal 5 opsi; template hanya mendukung opsi A-D.');
      const mp = ((getOne('mapel', T.m) || {}).obj || {}).nama || '';
      const gr = w.role === 'guru' ? w.g : (readT('guru').find(x => (x.m || []).indexOf(T.m) >= 0) || {});
      setByLabel(sI, 'Mata Pelajaran', mp); setByLabel(sI, 'Kelas', grp.nama); setByLabel(sI, 'Jenis Penilaian', T.nama);
      if (gr.nama) setByLabel(sI, 'Guru Mata Pelajaran', gr.nama);
      setByLabel(sI, 'NIP Guru', gr.nip || '-');
      if (c.kepala) setByLabel(sI, 'Kepala Madrasah', c.kepala);
      if (c.nipk) { sI.getRange('B16').setNumberFormat('@'); setByLabel(sI, 'NIP Kepala Madrasah', c.nipk); }
      sK.getRange('B6:C45').clearContent();
      sK.getRange(6, 2, ids.length, 2).setValues(ids.map(i => [L[sm[i].k], ((tm[sm[i].topik] || {}).nama || '') + ': ' + String(sm[i].q || '[soal bergambar]').slice(0, 90)]));
      const hs = {}; readT('hasil').forEach(h => { if (h.tes === T.id && h.fin) hs[h.p] = h; });
      const rows = readT('pes').filter(p => p.g === b.g && hs[p.id]).sort((a, c2) => String(a.nama).localeCompare(String(c2.nama))).slice(0, 50);
      sJ.getRange('B6:B55').clearContent(); sJ.getRange('D6:D55').clearContent(); sJ.getRange('E6:AR55').clearContent();
      if (rows.length) {
        sJ.getRange(6, 2, rows.length, 1).setValues(rows.map(p => [p.nama]));
        const rd = sJ.getRange(6, 4, rows.length, 1); rd.setNumberFormat('@'); rd.setValues(rows.map(p => [/^\d{8,}$/.test(String(p.u)) ? String(p.u) : '']));
        sJ.getRange(6, 5, rows.length, ids.length).setValues(rows.map(p => ids.map(i => { const a = hs[p.id].a[i]; return (a === undefined || a === null) ? '' : L[a]; })));
      }
      return ok({ n: rows.length, k: ids.length, link: ss.getUrl(), warn: warn });
    }
    case 'gprof': {
      const w = who(b.auth); if (!w || w.role !== 'guru') return err('Hanya untuk akun guru.');
      const r = getOne('guru', w.g.id); if (!r) return err('Guru tidak ditemukan.');
      const g = r.obj; g.nama = b.n || g.nama; g.nip = b.nip || ''; if (b.p) g.p = b.p;
      putOne('guru', r.row, g); return ok();
    }
    case 'pans': { // tanpa kunci global: hanya mengubah baris milik siswa itu
      const pes = pesFrom(b.tk); if (!pes) return err('Sesi tidak valid.');
      let bad = '';
      const r = updHasil(String(b.h || ''), h => {
        bad = '';
        if (h.p !== pes.id) { bad = 'Data tidak ditemukan.'; return false; }
        if (h.fin) return false;
        const T = tesById(h.tes); if (!T) { bad = 'Tes tidak ditemukan.'; return false; }
        if (Date.now() > endMs(h, T) + 30000) { bad = 'Waktu habis.'; return false; }
        h.a = clean(b.a, h);
      });
      if (r.miss) return err('Data tidak ditemukan.');
      if (r.busy) return err('Server sibuk, coba lagi.');
      return bad ? err(bad) : ok();
    }
    case 'pviol': {
      const pes = pesFrom(b.tk); if (!pes) return err('Sesi tidak valid.');
      let bad = '', res = null;
      const r = updHasil(String(b.h || ''), h => {
        bad = ''; res = null;
        if (h.p !== pes.id) { bad = 'Data tidak ditemukan.'; return false; }
        const T = tesById(h.tes); if (!T) { bad = 'Tes tidak ditemukan.'; return false; }
        if (h.fin) { res = { v: h.v || 0, fin: true, h: { fin: true, nilai: h.nilai, a: h.a } }; return false; }
        h.v = (h.v || 0) + 1;
        h.vl = (h.vl || []).concat([{ t: Date.now(), j: String(b.j || '').slice(0, 60) }]).slice(-20);
        const mx = T.anti === false ? 0 : (T.maxv === undefined ? 3 : (+T.maxv || 0));
        let fin = false;
        if (mx > 0 && h.v >= mx) {
          if (Date.now() <= endMs(h, T) + 60000 && b.a) h.a = clean(b.a, h);
          h.fin = true; h.nilai = score(h, T); fin = true;
        }
        res = { v: h.v, fin: fin, h: fin ? { fin: true, nilai: h.nilai, a: h.a } : undefined };
      });
      if (r.miss) return err('Data tidak ditemukan.');
      if (r.busy) return err('Server sibuk, coba lagi.');
      return bad ? err(bad) : ok(res);
    }
    case 'pfin': {
      const pes = pesFrom(b.tk); if (!pes) return err('Sesi tidak valid.');
      const id = String(b.h || ''); let bad = '', res = null;
      const apply = h => {
        bad = ''; res = null;
        if (h.p !== pes.id) { bad = 'Data tidak ditemukan.'; return false; }
        const T = tesById(h.tes); if (!T) { bad = 'Tes tidak ditemukan.'; return false; }
        if (h.fin) { res = { fin: true, nilai: h.nilai, a: h.a }; return false; }
        if (Date.now() <= endMs(h, T) + 60000) h.a = clean(b.a, h);
        h.fin = true; h.nilai = score(h, T);
        res = { fin: true, nilai: h.nilai, a: h.a };
      };
      const r = updHasil(id, apply);
      if (r.miss) return err('Data tidak ditemukan.');
      if (r.busy) return err('Server sibuk, coba lagi.');
      if (bad) return err(bad);
      if (r.ok) updHasil(id, h => (h.fin ? false : apply(h))); // verifikasi: pastikan tidak tertimpa autosave yang terlambat
      return ok({ h: res });
    }
  }
  return err('Aksi tidak dikenal.');
}
function clean(a, h) { const o = {}; h.ids.forEach(i => { if (a && a[i] !== undefined && a[i] !== null) o[i] = +a[i]; }); return o; }
