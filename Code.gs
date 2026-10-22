/**
 * API CBT Ujian Online MTsN 3 Subang
 * Google Sheets sebagai database. Deploy sebagai Web app:
 * Execute as: Me | Who has access: Anyone
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const TZ = SS.getSpreadsheetTimeZone();
const L5 = 'ABCDE';

/* ---------- utilitas ---------- */
const bool = v => v === true || String(v).toUpperCase() === 'TRUE';
const num = (v, d) => (v === '' || v == null || isNaN(+v)) ? d : +v;
const list = v => String(v == null ? '' : v).split(',').map(s => s.trim()).filter(Boolean);
const dt = v => v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm") : String(v == null ? '' : v);
const ms = v => v instanceof Date ? v.getTime() : Date.parse(v);
const pd = s => { try { return s ? Utilities.parseDate(s, TZ, "yyyy-MM-dd'T'HH:mm").getTime() : 0; } catch (e) { return 0; } };
const today = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const uid = () => Utilities.getUuid().replace(/-/g, '').slice(0, 7);
const shuf = a => { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const x = a[i]; a[i] = a[j]; a[j] = x; } return a; };
const ansStr = a => Object.keys(a || {}).map(k => k + ':' + L5[a[k]]).join(';');
const ansObj = s => { const a = {}; String(s || '').split(';').forEach(p => { const q = p.split(':'); if (q[0] && L5.indexOf(q[1]) >= 0) a[q[0]] = L5.indexOf(q[1]); }); return a; };
const parseJ = s => { try { return JSON.parse(s) || {}; } catch (e) { return {}; } };
const clean = (a, ids) => { const o = {}; ids.forEach(i => { const v = a && a[i]; if (Number.isInteger(v) && v >= 0 && v <= 4) o[i] = v; }); return o; };
const endT = (h, t) => { let e = h.mulai + (t.durasi + (h.x || 0)) * 60000; const s = pd(t.selesai); if (s) e = Math.min(e, s + (h.x || 0) * 60000); return e; };

/* ---------- skema tabel: kolom sheet <-> objek aplikasi ---------- */
const T = {
  Topik: {
    cols: ['id', 'nama', 'deskripsi'],
    toRow: o => [o.id, o.nama, o.desk || ''],
    fromRow: r => ({ id: String(r.id), nama: String(r.nama), desk: String(r.deskripsi) })
  },
  Soal: {
    cols: ['id', 'topik_id', 'pertanyaan', 'opsi_A', 'opsi_B', 'opsi_C', 'opsi_D', 'opsi_E', 'kunci'],
    toRow: o => [o.id, o.topik, o.q, o.o[0] || '', o.o[1] || '', o.o[2] || '', o.o[3] || '', o.o[4] || '', L5[o.k]],
    fromRow: r => ({
      id: String(r.id), topik: String(r.topik_id), q: String(r.pertanyaan),
      o: L5.split('').map(L => r['opsi_' + L]).filter(x => x !== '').map(String),
      k: L5.indexOf(String(r.kunci).toUpperCase())
    })
  },
  Group: {
    cols: ['id', 'nama'],
    toRow: o => [o.id, o.nama],
    fromRow: r => ({ id: String(r.id), nama: String(r.nama) })
  },
  Peserta: {
    cols: ['id', 'username', 'password', 'nama', 'group_id'],
    toRow: o => [o.id, o.u, o.p, o.nama, o.g],
    fromRow: r => ({ id: String(r.id), u: String(r.username), p: String(r.password), nama: String(r.nama), g: String(r.group_id) })
  },
  Tes: {
    cols: ['id', 'nama', 'mulai', 'selesai', 'durasi_menit', 'group_id', 'soal_id', 'poin_benar', 'poin_salah', 'poin_kosong', 'acak_soal', 'acak_jawaban', 'tampilkan_nilai', 'deskripsi', 'wajib_token'],
    toRow: o => [o.id, o.nama, o.mulai || '', o.selesai || '', o.durasi, (o.g || []).join(','), (o.ids || []).join(','), o.pb, o.ps, o.pk, !!o.as, !!o.aj, !!o.th, o.desk || '', !!o.token],
    fromRow: r => ({
      id: String(r.id), nama: String(r.nama), desk: String(r.deskripsi || ''), mulai: dt(r.mulai), selesai: dt(r.selesai),
      durasi: num(r.durasi_menit, 60), g: list(r.group_id), ids: list(r.soal_id),
      pb: num(r.poin_benar, 1), ps: num(r.poin_salah, 0), pk: num(r.poin_kosong, 0),
      as: bool(r.acak_soal), aj: bool(r.acak_jawaban), th: bool(r.tampilkan_nilai), token: bool(r.wajib_token)
    })
  },
  Hasil: {
    cols: ['id', 'tes_id', 'peserta_id', 'mulai', 'selesai', 'tambahan_waktu_menit', 'urutan_soal_id', 'jawaban', 'perm'],
    toRow: h => [h.id, h.tes, h.p, new Date(h.mulai).toISOString(), !!h.fin, h.x || 0, (h.ids || []).join(','), ansStr(h.a), JSON.stringify(h.perm || {})],
    fromRow: r => ({
      id: String(r.id), tes: String(r.tes_id), p: String(r.peserta_id), mulai: ms(r.mulai), fin: bool(r.selesai),
      x: num(r.tambahan_waktu_menit, 0), ids: list(r.urutan_soal_id), a: ansObj(r.jawaban), perm: parseJ(r.perm)
    })
  }
};

/* ---------- akses sheet ---------- */
function read(name) {
  const sh = SS.getSheetByName(name), lr = sh.getLastRow(), lc = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, lc).getValues()[0];
  if (lr < 2) return { sh: sh, h: head, rows: [] };
  const ids = sh.getRange(1, 1, lr, 1).getValues();
  let n = 0;
  for (let i = ids.length - 1; i >= 1; i--) if (ids[i][0] !== '') { n = i; break; }
  if (!n) return { sh: sh, h: head, rows: [] };
  const vals = sh.getRange(1, 1, n + 1, lc).getValues();
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === '') continue;
    const o = { _r: i + 1 };
    head.forEach((k, c) => { o[k] = vals[i][c]; });
    rows.push(o);
  }
  return { sh: sh, h: head, rows: rows };
}

function rowOf(sh, id) {
  const lr = sh.getLastRow();
  if (lr < 2) return 0;
  const f = sh.getRange(1, 1, lr, 1).createTextFinder(String(id)).matchEntireCell(true).matchCase(true).findNext();
  return f ? f.getRow() : 0;
}

function rowObj(name, id) {
  const sh = SS.getSheetByName(name), r = rowOf(sh, id);
  if (!r) return null;
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const v = sh.getRange(r, 1, 1, h.length).getValues()[0];
  const o = { _r: r };
  h.forEach((k, i) => { o[k] = v[i]; });
  return { sh: sh, h: h, r: r, o: o };
}

function writeRuns(sh, idx, rows, r0) {
  const need = r0 + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  let s = 0;
  for (let i = 1; i <= idx.length; i++) {
    if (i === idx.length || idx[i] !== idx[i - 1] + 1) {
      sh.getRange(r0, idx[s], rows.length, i - s).setValues(rows.map(v => v.slice(s, i)));
      s = i;
    }
  }
}

/** Tulis baris (array sesuai cols; elemen pertama = kunci). Kolom rumus tidak disentuh. */
function put(name, cols, rows, onlyNew) {
  const t = read(name), sh = t.sh, h = t.h.slice();
  const idx = cols.map(c => {
    let i = h.indexOf(c);
    if (i < 0) { h.push(c); i = h.length - 1; sh.getRange(1, i + 1).setValue(c); }
    return i + 1;
  });
  const pos = {};
  t.rows.forEach(r => { pos[String(r[t.h[0]])] = r._r; });
  let next = 2;
  t.rows.forEach(r => { if (r._r >= next) next = r._r + 1; });
  const add = [];
  rows.forEach(v => {
    const r = pos[String(v[0])];
    if (r) { if (!onlyNew) writeRuns(sh, idx, [v], r); } else add.push(v);
  });
  if (add.length) writeRuns(sh, idx, add, next);
}

function del(name, ids) {
  const t = read(name), set = {};
  ids.forEach(i => { set[String(i)] = 1; });
  t.rows.filter(r => set[String(r[t.h[0]])]).map(r => r._r).sort((a, b) => b - a).forEach(r => t.sh.deleteRow(r));
}

const withLock = fn => { const l = LockService.getScriptLock(); l.waitLock(30000); try { return fn(); } finally { l.releaseLock(); } };

/* ---------- konfigurasi, token, sesi ---------- */
function cfg() {
  const o = {};
  read('Konfigurasi').rows.forEach(r => { o[r.kunci] = String(r.nilai); });
  return { nama: o.nama || 'CBT', ket: o.ket || '', au: o.admin_user || '', ap: o.admin_pass || '' }; // ap = hash/teks tersimpan; jangan kirim ke klien
}
const pub = () => { const c = cfg(); return { nama: c.nama, ket: c.ket }; };
/* ---------- keamanan: hash password admin & pembatasan percobaan login ---------- */
const hex = bytes => bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
const hashPw = (p, salt) => 'h1$' + salt + '$' + hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + p, Utilities.Charset.UTF_8));
const newHash = p => hashPw(p, uid() + uid());
function samePw(stored, p) {
  stored = String(stored || ''); p = String(p == null ? '' : p);
  if (stored.indexOf('h1$') === 0) return hashPw(p, stored.split('$')[1]) === stored;
  return stored !== '' && stored === p;
}
const setAdminPw = v => put('Konfigurasi', ['kunci', 'nilai'], [['admin_pass', v]]);
function adminOK(u, p) {
  const c = cfg();
  if (!c.au || String(u) !== c.au || !samePw(c.ap, p)) return false;
  if (String(c.ap).indexOf('h1$') !== 0) setAdminPw(newHash(String(p))); // migrasi otomatis dari teks biasa
  return true;
}
const isAdmin = a => !!(a && adminOK(a.u, a.p));

const LOCK_SEC = 600;
const lkKey = k => 'f_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(k))).slice(0, 24);
const locked = (k, max) => num(CacheService.getScriptCache().get(lkKey(k)), 0) >= max;
const failed = k => { const c = CacheService.getScriptCache(); c.put(lkKey(k), String(num(c.get(lkKey(k)), 0) + 1), LOCK_SEC); };
const passed = k => CacheService.getScriptCache().remove(lkKey(k));
const LOCKMSG = 'Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.';
const weakPw = p => String(p).length < 6 || ['admin', 'password', '123456', '12345678', 'admin123'].indexOf(String(p).toLowerCase()) >= 0;

function tokenNow() {
  const v = SS.getSheetByName('Token').getRange(2, 1, 1, 2).getValues()[0];
  const d = v[1] instanceof Date ? Utilities.formatDate(v[1], TZ, 'yyyy-MM-dd') : String(v[1] || '');
  return { c: String(v[0] || ''), d: d };
}

function secret() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('S');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('S', s); }
  return s;
}
const sig = id => Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(id), secret())).slice(0, 22);
const tok = id => id + '.' + sig(id);
function chk(tk) {
  tk = String(tk || '');
  const i = tk.lastIndexOf('.');
  if (i < 1) return '';
  const id = tk.slice(0, i);
  return sig(id) === tk.slice(i + 1) ? id : '';
}

/* ---------- data tes (di-cache 2 menit) ---------- */
function bundle(tid) {
  const c = CacheService.getScriptCache(), key = 'b' + (c.get('v') || '0') + '_' + tid;
  const hit = c.get(key);
  if (hit) return JSON.parse(hit);
  const tr = rowObj('Tes', tid);
  if (!tr) return {};
  const t = T.Tes.fromRow(tr.o), want = {}, soal = {};
  t.ids.forEach(i => { want[i] = 1; });
  read('Soal').rows.forEach(r => { if (want[String(r.id)]) soal[String(r.id)] = T.Soal.fromRow(r); });
  const out = { t: t, soal: soal };
  try { c.put(key, JSON.stringify(out), 120); } catch (e) { /* terlalu besar: abaikan */ }
  return out;
}
const bust = () => CacheService.getScriptCache().put('v', String(Date.now()), 21600);

function grade(h, t, soal) {
  let pts = 0;
  h.ids.forEach(i => {
    const s = soal[i];
    if (!s) return;
    const a = h.a[i];
    pts += a === undefined ? t.pk : a === s.k ? t.pb : t.ps;
  });
  const mx = h.ids.length * t.pb;
  return mx > 0 ? Math.round(Math.max(0, pts) / mx * 10000) / 100 : pts;
}

function findH(pid, tid) {
  const r = read('Hasil').rows.find(x => String(x.peserta_id) === String(pid) && String(x.tes_id) === String(tid));
  return r ? T.Hasil.fromRow(r) : null;
}

function newH(pid, t, soal) {
  const ids = (t.as ? shuf(t.ids) : t.ids.slice()).filter(i => soal[i]);
  const perm = {};
  ids.forEach(i => {
    const ix = soal[i].o.map((_, n) => n);
    perm[i] = t.aj ? shuf(ix) : ix;
  });
  const h = { id: uid(), tes: t.id, p: pid, mulai: Date.now(), ids: ids, perm: perm, a: {}, fin: false, x: 0 };
  put('Hasil', T.Hasil.cols, [T.Hasil.toRow(h)], true);
  return h;
}

/* ---------- aksi ---------- */
const slim = t => ({ id: t.id, nama: t.nama, desk: t.desk, mulai: t.mulai, selesai: t.selesai, durasi: t.durasi, g: t.g, ids: t.ids, token: t.token, th: t.th });
const NO = err => ({ ok: false, err: err });

const A = {
  info: () => ({ ok: true, cfg: pub() }),

  alogin: b => {
    if (locked('admin', 5)) return NO(LOCKMSG);
    if (!adminOK(b.u, b.p)) { failed('admin'); return NO('Login admin salah.'); }
    passed('admin');
    const c = cfg();
    c.ap = ''; // password tidak pernah dikirim balik ke browser
    const g = n => read(n).rows;
    const soal = g('Soal').map(T.Soal.fromRow), sm = {};
    soal.forEach(s => { sm[s.id] = s; });
    const tes = g('Tes').map(T.Tes.fromRow), tm = {};
    tes.forEach(t => { tm[t.id] = t; });
    const hasil = g('Hasil').map(T.Hasil.fromRow);
    hasil.forEach(h => { if (h.fin && tm[h.tes]) h.nilai = grade(h, tm[h.tes], sm); });
    const tk = tokenNow();
    return { ok: true, S: { cfg: c, topik: g('Topik').map(T.Topik.fromRow), soal: soal, grp: g('Group').map(T.Group.fromRow), pes: g('Peserta').map(T.Peserta.fromRow), tes: tes, token: { c: tk.c, d: tk.d }, hasil: hasil, warn: weakPw(b.p) ? 'Password admin Anda lemah/bawaan. Segera ganti di Pengaturan & Backup > Ubah identitas & akun admin.' : '' } };
  },

  sync: b => {
    if (!isAdmin(b.auth)) return NO('Akses ditolak.');
    const ch = b.ch || {};
    return withLock(() => {
      const map = { topik: 'Topik', soal: 'Soal', grp: 'Group', pes: 'Peserta', tes: 'Tes' };
      Object.keys(map).forEach(k => {
        const c = ch[k];
        if (!c) return;
        const sp = T[map[k]];
        if (c.up.length) put(map[k], sp.cols, c.up.map(sp.toRow));
        if (c.rm.length) del(map[k], c.rm);
      });
      if (ch.hasil) {
        const sp = T.Hasil, t = read('Hasil');
        const cx = t.h.indexOf('tambahan_waktu_menit') + 1, cf = t.h.indexOf('selesai') + 1, pos = {}, add = [];
        t.rows.forEach(r => { pos[String(r.id)] = r; });
        ch.hasil.up.forEach(h => {
          const r = pos[String(h.id)];
          if (!r) { add.push(sp.toRow(h)); return; }
          if (cx > 0) t.sh.getRange(r._r, cx).setValue(h.x || 0);
          if (h.fin && !bool(r.selesai) && cf > 0) t.sh.getRange(r._r, cf).setValue(true);
        });
        if (add.length) put('Hasil', sp.cols, add, true);
        if (ch.hasil.rm.length) del('Hasil', ch.hasil.rm);
      }
      if (ch.cfg) put('Konfigurasi', ['kunci', 'nilai'], [['nama', ch.cfg.nama], ['ket', ch.cfg.ket], ['admin_user', ch.cfg.au], ['admin_pass', ch.cfg.ap ? newHash(String(ch.cfg.ap)) : cfg().ap]]); // kosong = password tidak diubah
      if (ch.token) SS.getSheetByName('Token').getRange(2, 1, 1, 2).setNumberFormat('@').setValues([[String(ch.token.c || ''), ch.token.c ? today() : '']]);
      bust();
      return { ok: true };
    });
  },

  plogin: b => {
    const lk = 'p:' + String(b.u).trim().toLowerCase();
    if (locked(lk, 8)) return NO(LOCKMSG);
    const p = read('Peserta').rows.find(r => String(r.username) === String(b.u).trim() && String(r.password) === String(b.p));
    if (!p) { failed(lk); return NO('Username/password peserta salah.'); }
    passed(lk);
    const gid = String(p.group_id);
    const grp = read('Group').rows.filter(r => String(r.id) === gid).map(T.Group.fromRow);
    const tes = read('Tes').rows.map(T.Tes.fromRow).filter(t => t.g.indexOf(gid) >= 0);
    const tm = {};
    tes.forEach(t => { tm[t.id] = t; });
    const hasil = read('Hasil').rows.filter(r => String(r.peserta_id) === String(p.id)).map(T.Hasil.fromRow)
      .filter(h => tm[h.tes]).map(h => {
        const o = { id: h.id, tes: h.tes, p: h.p, mulai: h.mulai, ids: h.ids, perm: h.perm, a: h.a, fin: h.fin, x: h.x };
        if (h.fin && tm[h.tes].th) o.nilai = grade(h, tm[h.tes], bundle(h.tes).soal || {});
        return o;
      });
    return { ok: true, tk: tok(String(p.id)), cfg: pub(), pes: { id: String(p.id), u: String(p.username), nama: String(p.nama), g: gid }, grp: grp, tes: tes.map(slim), hasil: hasil };
  },

  pstart: b => {
    const pid = chk(b.tk);
    if (!pid) return NO('Sesi berakhir, silakan login ulang.');
    const bn = bundle(b.tes), t = bn.t;
    if (!t) return NO('Tes tidak ditemukan.');
    const po = rowObj('Peserta', pid);
    if (!po || t.g.indexOf(String(po.o.group_id)) < 0) return NO('Tes ini bukan untuk Anda.');
    let h = findH(pid, t.id);
    if (!h) {
      const now = Date.now(), s = pd(t.mulai), e = pd(t.selesai);
      if ((s && now < s) || (e && now > e)) return NO('Tes belum dibuka atau sudah ditutup.');
      if (t.token) {
        const k = tokenNow();
        if (!b.token || !k.c || k.d !== today() || String(b.token).trim().toUpperCase() !== k.c.toUpperCase()) return NO('Token salah atau belum dibuat.');
      }
      if (!t.ids.length) return NO('Soal belum tersedia.');
      h = withLock(() => findH(pid, t.id) || newH(pid, t, bn.soal));
    } else if (h.fin) return NO('Tes sudah selesai.');
    const soal = h.ids.map(i => bn.soal[i]).filter(Boolean).map(s => ({ id: s.id, q: s.q, o: s.o }));
    return { ok: true, h: { id: h.id, tes: h.tes, p: h.p, mulai: h.mulai, ids: h.ids, perm: h.perm, a: h.a, fin: false, x: h.x }, soal: soal };
  },

  pans: b => {
    const pid = chk(b.tk);
    if (!pid) return NO('Sesi berakhir.');
    const ro = rowObj('Hasil', b.h);
    if (!ro || String(ro.o.peserta_id) !== pid) return NO('Data tidak ditemukan.');
    const h = T.Hasil.fromRow(ro.o);
    if (h.fin) return { ok: true, fin: true };
    const t = bundle(h.tes).t;
    if (t && Date.now() > endT(h, t) + 60000) return { ok: false, err: 'Waktu habis.', over: true };
    ro.sh.getRange(ro.r, ro.h.indexOf('jawaban') + 1).setValue(ansStr(clean(b.a, h.ids)));
    return { ok: true };
  },

  pfin: b => {
    const pid = chk(b.tk);
    if (!pid) return NO('Sesi berakhir.');
    const ro = rowObj('Hasil', b.h);
    if (!ro || String(ro.o.peserta_id) !== pid) return NO('Data tidak ditemukan.');
    const h = T.Hasil.fromRow(ro.o), bn = bundle(h.tes), t = bn.t;
    if (!t) return NO('Tes tidak ditemukan.');
    if (!h.fin) {
      if (b.a && Date.now() <= endT(h, t) + 60000) h.a = clean(b.a, h.ids);
      ro.sh.getRange(ro.r, ro.h.indexOf('jawaban') + 1).setValue(ansStr(h.a));
      ro.sh.getRange(ro.r, ro.h.indexOf('selesai') + 1).setValue(true);
      h.fin = true;
    }
    const out = { id: h.id, fin: true, a: h.a };
    if (t.th) out.nilai = grade(h, t, bn.soal);
    return { ok: true, h: out };
  }
};

/* ---------- pintu masuk ---------- */
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() { return ContentService.createTextOutput('API CBT aktif.'); }

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out(NO('Permintaan tidak valid.')); }
  const f = A[b.action];
  if (!f) return out(NO('Aksi tidak dikenal.'));
  try { return out(f(b)); } catch (x) {
    const m = String(x && x.message || x);
    return out(NO(/lock/i.test(m) ? 'Server sibuk, coba lagi.' : 'Kesalahan server: ' + m));
  }
}
