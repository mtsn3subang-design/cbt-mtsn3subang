/**
 * Penghubung CBT -> Aplikasi Analisis Butir Soal
 * Tambahkan file ini ke project Apps Script yang sama dengan Code.gs.
 * Menu baru "CBT" muncul di Google Sheets: CBT > Ekspor ke Analisis Butir.
 * Hasilnya dua sheet (Ekspor_KUNCI dan Ekspor_JAWABAN) yang tata letaknya sama
 * dengan sheet KUNCI dan JAWABAN di file Analisis Butir. Tinggal salin-tempel.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('CBT')
    .addItem('Ekspor ke Analisis Butir…', 'menuExportAnalisis')
    .addToUi();
}

/** Fungsi murni (tanpa akses Sheets) agar mudah diuji. */
function buildAnalisis(t, soal, peserta, hasil, maxItem, maxSiswa) {
  maxItem = maxItem || 40; maxSiswa = maxSiswa || 50;
  const warn = [];
  let ids = t.ids.filter(i => soal[i]);
  if (ids.length < t.ids.length) warn.push((t.ids.length - ids.length) + ' soal pada tes sudah dihapus dari bank soal dan dilewati.');
  if (ids.length > maxItem) { warn.push('Soal ' + ids.length + ' butir, template hanya memuat ' + maxItem + '. Butir ke-' + (maxItem + 1) + ' dst dipotong.'); ids = ids.slice(0, maxItem); }
  const kunci = ids.map(i => 'ABCDE'.charAt(soal[i].k));
  const pm = {};
  peserta.forEach(p => { pm[p.id] = p; });
  let rows = hasil.filter(h => h.tes === t.id && h.fin).map(h => {
    const p = pm[h.p] || { nama: '(peserta ' + h.p + ')', u: '' };
    return { nama: p.nama, nisn: p.u, a: ids.map(i => h.a[i] === undefined ? '' : 'ABCDE'.charAt(h.a[i])) };
  }).sort((x, y) => x.nama.localeCompare(y.nama));
  const belum = hasil.filter(h => h.tes === t.id && !h.fin).length;
  if (belum) warn.push(belum + ' peserta belum menyelesaikan tes dan tidak diikutkan.');
  if (rows.length > maxSiswa) { warn.push('Peserta ' + rows.length + ' orang, template hanya memuat ' + maxSiswa + '. Pisahkan per kelas.'); rows = rows.slice(0, maxSiswa); }
  const adaE = kunci.some(k => k === 'E') || rows.some(r => r.a.indexOf('E') >= 0);
  if (adaE) warn.push('Ada opsi E. Template Analisis Butir hanya mengenal A-D; butir tersebut akan terbaca salah.');
  return { ids: ids, kunci: kunci, rows: rows, warn: warn };
}

function exportAnalisis(tesId) {
  const tr = read('Tes').rows.find(r => String(r.id) === String(tesId));
  if (!tr) throw new Error('ID tes tidak ditemukan: ' + tesId);
  const t = T.Tes.fromRow(tr), soal = {};
  read('Soal').rows.forEach(r => { soal[String(r.id)] = T.Soal.fromRow(r); });
  const gm = {};
  read('Group').rows.forEach(r => { gm[String(r.id)] = String(r.nama); });
  const peserta = read('Peserta').rows.map(T.Peserta.fromRow);
  peserta.forEach(p => { p.grup = gm[p.g] || ''; });
  const hasil = read('Hasil').rows.map(T.Hasil.fromRow);
  const R = buildAnalisis(t, soal, peserta, hasil);

  const sheet = name => { let sh = SS.getSheetByName(name); if (sh) sh.clear(); else sh = SS.insertSheet(name); return sh; };
  const k = sheet('Ekspor_KUNCI');
  k.getRange(1, 1).setValue('Salin B6 ke bawah ke sheet KUNCI (B6) pada file Analisis Butir. Tes: ' + t.nama);
  k.getRange(5, 1, 1, 2).setValues([['No', 'Kunci Jawaban (A/B/C/D)']]);
  if (R.kunci.length) k.getRange(6, 1, R.kunci.length, 2).setValues(R.kunci.map((x, i) => [i + 1, x]));

  const j = sheet('Ekspor_JAWABAN');
  j.getRange(1, 1).setValue('Salin B6 sampai akhir ke sheet JAWABAN (B6) pada file Analisis Butir. Kolom C = nama group, D = username (isi NISN sebagai username agar terbawa).');
  const head = ['No', 'Nama Siswa', 'Kelas', 'NISN'].concat(R.kunci.map((_, i) => i + 1));
  j.getRange(5, 1, 1, head.length).setValues([head]);
  if (R.rows.length) {
    const grup = {};
    peserta.forEach(p => { grup[p.nama] = p.grup; });
    j.getRange(6, 1, R.rows.length, head.length).setValues(R.rows.map((r, i) => [i + 1, r.nama, grup[r.nama] || '', "'" + r.nisn].concat(r.a)));
  }
  j.getRange(6, 4, Math.max(R.rows.length, 1), 1).setNumberFormat('@');
  return { butir: R.kunci.length, siswa: R.rows.length, warn: R.warn };
}

function menuExportAnalisis() {
  const ui = SpreadsheetApp.getUi();
  const daftar = read('Tes').rows.map(r => r.id + ' = ' + r.nama).join('\n');
  const res = ui.prompt('Ekspor ke Analisis Butir', 'Ketik ID tes:\n\n' + daftar, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  try {
    const r = exportAnalisis(res.getResponseText().trim());
    ui.alert('Selesai', r.butir + ' butir, ' + r.siswa + ' siswa.\nSheet Ekspor_KUNCI dan Ekspor_JAWABAN sudah dibuat.' + (r.warn.length ? '\n\nPERHATIAN:\n- ' + r.warn.join('\n- ') : ''), ui.ButtonSet.OK);
  } catch (e) { ui.alert('Gagal', String(e.message || e), ui.ButtonSet.OK); }
}

if (typeof module !== 'undefined') module.exports = { buildAnalisis: buildAnalisis };
