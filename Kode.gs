/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ROAD TO GRAND MASTER · PANTAU — Pencatat Harian (Google Apps Script)
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * PERAN SKRIP INI — dan yang BUKAN perannya.
 *
 * Perannya ada satu: menjadi SAKSI yang tidak pernah lupa.
 * Lichess hanya memberi keadaan SAAT INI lewat API. Kalau tidak ada yang
 * memotretnya setiap hari, hari itu hilang selamanya. Browser tidak bisa memotret saat tidak
 * dibuka. Skrip ini bisa — trigger harian jam 23:00, tanpa manusia.
 *
 * Yang BUKAN perannya: menjadi database. Sheet di sini adalah buku catatan,
 * bukan Postgres. Jangan tambahkan tabel relasional ke sini.
 *
 * ── CARA PASANG ──────────────────────────────────────────────────────────
 * 1. Buat Google Spreadsheet baru.
 * 2. Extensions > Apps Script. Hapus isi Code.gs, tempel seluruh file ini.
 * 3. Ubah KONFIG di bawah (username, email).
 * 4. Jalankan fungsi  pasang()  sekali. Izinkan akses saat diminta.
 * 5. Selesai. Skrip berjalan sendiri tiap malam.
 *
 * ── RISIKO YANG SUDAH DIKETAHUI ─────────────────────────────────────────
 * Permintaan ke API kadang gagal (jaringan / rate limit). Kalau itu terjadi,
 * skrip ini TIDAK akan
 * menulis angka nol seolah anak tidak latihan — ia menulis "GAGAL" di kolom
 * status dan mengirim email peringatan. Kesalahan diam adalah kesalahan
 * terburuk untuk aplikasi pengawasan.
 */

// ═══════════════════ KONFIG ═══════════════════
const KONFIG = {
  /* Banyak anak. Tambahkan objek baru ke daftar ini; tidak perlu apa pun lagi.
     Setiap anak mendapat barisnya sendiri di sheet, dibedakan kolom "Anak". */
  anak: [
    // GANTI nama & username di bawah dengan data anak Anda sendiri di SALINAN
    // PRIBADI. Kalau file ini akan di-commit ke repo publik, JANGAN isi nama
    // asli / username asli di sini — nama anak dan akun game-nya jadi tertaut
    // publik selamanya (riwayat Git tidak melupakan).
    { nama: 'NamaAnak', lichess: 'UsernameLichess' },
    // { nama: 'Adik',  lichess: '' },
  ],

  // Email penerima laporan harian. Kosongkan array untuk mematikan email.
  emailOrangTua: ['ganti@email-anda.com'],
  emailPelatih:  [],

  /* Telegram (opsional). Laporan malam langsung masuk ke genggaman, tidak
     menunggu email dibuka. Gratis dan resmi — berbeda dengan gateway WhatsApp
     tak resmi, yang berbayar atau melanggar ToS WhatsApp dan bisa memblokir
     nomormu. Kalau ingin di WhatsApp: teruskan dari Telegram, jangan bot bajakan.

     Cara: chat @BotFather -> /newbot -> salin token.
           Chat bot itu sekali, lalu buka
           https://api.telegram.org/bot<TOKEN>/getUpdates -> salin "chat":{"id":...}. */
  telegramToken:  '',
  telegramChatId: '',

  // Kontak untuk header User-Agent — sopan santun terhadap API publik.
  kontak: 'ganti@email-anda.com',

  // Target harian. Harus sama dengan yang di dasbor HTML.
  target: {
    puzzle:         30,  // puzzle per hari di Lichess — SATU-SATUNYA hitungan puzzle
    rapid:           2,  // partai rapid per hari         yang bisa diverifikasi API
    bulletMaks:      0,  // batas partai bullet (pagar, bukan target)
    partaiMaks:      6,  // batas total partai per hari (anti-tilt)
  },

  jamCatat: 23,          // jam trigger harian (0-23), waktu spreadsheet
};

const SHEET_HARIAN = 'Harian';
const SHEET_PARTAI = 'Partai';
const SHEET_LOG    = 'Log';

// ═══════════════════ PEMASANGAN ═══════════════════
function pasang() {
  siapkanSheet_();
  hapusTriggerLama_();
  ScriptApp.newTrigger('catatHarian')
    .timeBased().atHour(KONFIG.jamCatat).everyDays(1).create();
  catatHarian();                        // jalankan sekali sekarang
  SpreadsheetApp.getUi && SpreadsheetApp.getActive().toast(
    KONFIG.anak.length + ' anak terpasang. Skrip akan berjalan tiap hari jam ' + KONFIG.jamCatat + ':00.', 'RTGM Pantau', 8);
}

function hapusTriggerLama_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['catatHarian'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
}

function siapkanSheet_() {
  const ss = SpreadsheetApp.getActive();
  const shH = mk_(ss, SHEET_HARIAN, [
    'Tanggal', 'Anak', 'Status ambil',
    'Puzzle (Lichess)', 'Puzzle benar', 'Puzzle salah', 'Rating puzzle',
    'Puzzle TOTAL kumulatif',
    'Partai hari ini', 'Rapid', 'Blitz', 'Bullet', 'Menang', 'Kalah', 'Seri',
    'Kalah beruntun', 'Akurasi rata2', 'Kepatuhan %', 'Pelanggaran',
    'Tema (JSON)',
  ]);
  // Kolom tanggal dipaksa berformat TEKS. Tanpa ini, Sheets terus-menerus
  // mengubahnya jadi Date dan perbandingan string jadi rapuh.
  shH.getRange('A:A').setNumberFormat('@');

  mk_(ss, SHEET_PARTAI, [
    'Anak', 'Waktu selesai', 'Kontrol', 'Lawan', 'Rating lawan', 'Hasil',
    'Rating anak', 'Akurasi anak', 'ECO', 'URL',
  ]);
  mk_(ss, SHEET_LOG, ['Waktu', 'Tingkat', 'Pesan']);
}

/* BUG YANG DIPERBAIKI: Sheets otomatis mengubah string "2026-07-14" menjadi
   objek Date. getValues() lalu mengembalikan Date, dan `Date === "2026-07-14"`
   SELALU false. Akibatnya baris hari ini tidak pernah ditimpa — setiap hari
   bertambah baris duplikat, dan grafik ikut kacau.

   getDisplayValues() BUKAN solusinya: ia mengembalikan tanggal sesuai format
   tampilan Sheet, yang di lokal Indonesia berbunyi "14/07/2026". Itu menukar
   satu bug dengan bug yang lebih sunyi.

   Yang benar: normalkan KEDUA sisi, apa pun bentuk yang dikembalikan Sheets. */
function tglStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var t = String(v == null ? '' : v).trim();
  var m = t.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);   // 14/07/2026
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return t;
}

function mk_(ss, nama, header) {
  let sh = ss.getSheetByName(nama);
  if (!sh) sh = ss.insertSheet(nama);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#1B2130').setFontColor('#F1F2EB');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ═══════════════════ PENGAMBILAN DATA ═══════════════════
/**
 * Semua permintaan lewat sini. Tidak ada pemanggilan UrlFetchApp langsung
 * di tempat lain — supaya penanganan 403 Cloudflare hanya ada di satu tempat.
 */
function ambil_(url) {
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'RTGM-Pantau/1.0 (kontak: ' + KONFIG.kontak + ')',
      'Accept': 'application/json',
    },
  });
  const kode = res.getResponseCode();
  if (kode === 403) throw new Error('403 — permintaan ditolak server.');
  if (kode === 404) throw new Error('404 — username tidak ditemukan: ' + ANAK.lichess);
  if (kode === 429) throw new Error('429 — terlalu banyak permintaan. Kurangi frekuensi trigger.');
  if (kode !== 200) throw new Error('HTTP ' + kode + ' dari ' + url);
  return JSON.parse(res.getContentText());
}

/* Anak yang sedang diproses. Diset oleh catatHarian() sebelum tiap putaran. */
let ANAK = KONFIG.anak[0];
const LI   = () => 'https://lichess.org/api/user/' + (ANAK.lichess || '');

/**
 * Lichess: SATU-SATUNYA sumber jumlah puzzle harian yang bisa diverifikasi.
 * PENTING: timestamp Lichess dalam MILIDETIK.
 * Menyamakan keduanya adalah bug klasik yang membuat angka puzzle jadi nol
 * tanpa error apa pun — dan orang tua menghukum anak atas kesalahan kode.
 */
function ambilLichess_(tanggal, tz) {
  const hasil = { puzzle: 0, benar: 0, salah: 0, ratingPuzzle: '', puzzleTotal: null,
                  partai: 0, partaiFeed: 0, partaiList: [], status: 'OK' };
  if (!ANAK.lichess) { hasil.status = 'MATI'; return hasil; }
  try {
    const prof = ambil_(LI_URL_());
    if (prof && prof.perfs && prof.perfs.puzzle) hasil.ratingPuzzle = prof.perfs.puzzle.rating;

    if (prof && prof.perfs && prof.perfs.puzzle) hasil.puzzleTotal = prof.perfs.puzzle.games || 0;

    const act = ambil_(LI() + '/activity');
    (act || []).forEach(a => {
      if (!a.interval) return;
      // BUG YANG DIPERBAIKI: interval Lichess sering BERAKHIR di tengah malam
      // berikutnya. Mencocokkan pada `end` melempar aktivitas hari ini ke hari
      // BESOK, sehingga hari ini tampak kosong. Cocokkan pada hari MULAInya.
      const mulai = new Date(msLi_(a.interval.start));
      if (Utilities.formatDate(mulai, tz, 'yyyy-MM-dd') !== tanggal) return;
      if (a.puzzles && a.puzzles.score) {
        const sc = a.puzzles.score;
        hasil.benar = sc.win || 0;
        hasil.salah = sc.loss || 0;
        hasil.puzzle = hasil.benar + hasil.salah + (sc.draw || 0);
      }
      // Feed aktivitas dipakai sebagai CADANGAN saja untuk partai.
      if (a.games) {
        Object.keys(a.games).forEach(k => {
          const v = a.games[k];
          hasil.partaiFeed += (v.win || 0) + (v.loss || 0) + (v.draw || 0);
        });
      }
    });

    /* Partai sungguhan, bukan ringkasan. Lichess mengirim NDJSON:
       satu objek JSON per baris. JSON.parse() atas seluruh badan akan GAGAL. */
    /* BUG FATAL YANG DIPERBAIKI: dulu ditulis LI() + '/games/user/' + nama,
       padahal LI() sudah berisi '/api/user/{nama}'. Hasilnya:
         lichess.org/api/user/X/games/user/X   -> 404, SELALU.
       Endpoint ekspor partai Lichess ada di cabang yang BERBEDA:
         lichess.org/api/games/user/{nama}
       Akibat bug ini, daftar partai Lichess tidak pernah sekali pun berhasil
       ditarik, dan skrip diam-diam mundur ke feed aktivitas yang kurang akurat. */
    const urlPartai = 'https://lichess.org/api/games/user/' + encodeURIComponent(ANAK.lichess) +
      '?max=100&opening=true';
    const raw = UrlFetchApp.fetch(urlPartai, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'Accept': 'application/x-ndjson',
        'User-Agent': 'RoadToGrandMaster/1.0 (kontak: ' + KONFIG.kontak + ')',
      },
    });
    if (raw.getResponseCode() !== 200)
      log_('PERINGATAN', 'Lichess partai HTTP ' + raw.getResponseCode() + ' — ' + urlPartai);
    if (raw.getResponseCode() === 200) {
      const me = ANAK.lichess.toLowerCase();
      raw.getContentText().split('\n').filter(String).forEach(baris => {
        let g; try { g = JSON.parse(baris); } catch (e) { return; }
        const t = new Date(msLi_(g.lastMoveAt || g.createdAt));
        if (Utilities.formatDate(t, tz, 'yyyy-MM-dd') !== tanggal) return;
        const w = g.players && g.players.white, b = g.players && g.players.black;
        if (!w || !b) return;
        const idW = ((w.user && (w.user.id || w.user.name)) || '').toLowerCase();
        const isW = idW === me;
        let speed = g.speed === 'ultraBullet' ? 'bullet' : g.speed;
        let res = 'R';
        if (g.winner) res = (g.winner === (isW ? 'white' : 'black')) ? 'M' : 'K';
        hasil.partaiList.push({ speed: speed, res: res, t: t.getTime() });
      });
      hasil.partai = hasil.partaiList.length;
    } else {
      // gagal tarik daftar partai -> pakai angka dari feed
      hasil.partai = hasil.partaiFeed;
      hasil.status = 'OK (partai dari feed, daftar partai gagal)';
    }
  } catch (e) {
    hasil.status = 'GAGAL: ' + e.message;
    log_('PERINGATAN', 'Lichess: ' + e.message);
  }

  /* JARING PENGAMAN: kalau feed aktivitas melaporkan 0 tapi penghitung
     kumulatif naik dibanding baris kemarin, yang benar adalah selisihnya.
     Feed Lichess bisa terlambat; penghitung total tidak pernah bohong. */
  if (hasil.puzzle === 0 && hasil.puzzleTotal != null) {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HARIAN);
    const data = sh.getDataRange().getValues();
    let sebelum = null;
    for (let i = data.length - 1; i >= 1; i--) {
      const d = tglStr_(data[i][0], tz);
      if (d && d < tanggal && data[i][1] === ANAK.nama &&
          data[i][7] !== '' && data[i][7] != null) {
        sebelum = Number(data[i][7]); break;
      }
    }
    if (sebelum != null && hasil.puzzleTotal > sebelum) {
      hasil.puzzle = hasil.puzzleTotal - sebelum;
      hasil.status = 'OK (dari penghitung kumulatif, feed kosong)';
    }
  }
  return hasil;
}
function LI_URL_() { return LI(); }
/** Lichess kirim milidetik. Kalau angkanya kecil, itu detik — jangan asal kali 1000. */
function msLi_(t) { return t < 1e11 ? t * 1000 : t; }

// ═══════════════════ TOKEN LICHESS — verifikasi TEMA puzzle ═══════════════════
/* Token puzzle:read TIDAK ditulis di file ini. Simpan di Script Properties:
   Project Settings > Script properties > tambah  LICHESS_TOKEN = <token puzzle:read>.
   (Atau tempel di pasangTokenLichess() sekali, jalankan, lalu kosongkan lagi.)
   /api/puzzle/activity selalu mengembalikan aktivitas PEMILIK token — jadi token
   ini harus milik akun Lichess si anak. Banyak anak = butuh mekanisme
   token per anak; untuk satu anak, ini sudah benar. */
/* Opsi "tempel di kode": isi token puzzle:read di antara kutip di bawah, sekali.
   Berlaku di semua perangkat, tak perlu tempel di browser lagi.
   PERINGATAN: JANGAN pakai opsi ini kalau file ini di-commit ke Git / dibagikan.
   Token yang ditempel di sini ikut ke riwayat Git SELAMANYA, bahkan setelah
   dihapus lagi — commit lama masih menyimpannya. Kalau repo ini publik (atau
   bisa jadi publik), token akan bocor. Pakai Script Properties (lihat
   pasangTokenLichess() di bawah) yang TIDAK pernah masuk kontrol versi. */
const LICHESS_TOKEN_HARDCODE = '';
function _tokenLi_() {
  return PropertiesService.getScriptProperties().getProperty('LICHESS_TOKEN')
      || (LICHESS_TOKEN_HARDCODE || '').trim();
}
function pasangTokenLichess() {
  const t = '';                 // <- tempel token puzzle:read di sini, jalankan SEKALI
  if (!t) { console.log('Isi variabel t di dalam fungsi ini dulu, lalu jalankan.'); return; }
  PropertiesService.getScriptProperties().setProperty('LICHESS_TOKEN', t.trim());
  console.log('Token tersimpan di Script Properties. SEKARANG kosongkan lagi variabel t di kode ini.');
}
/* ═════════ MIGRASI: HAPUS KOLOM CHESS.COM DARI SHEET (sekali jalan) ═════════
   WAJIB dijalankan SETELAH menempel kode ini, SEBELUM catatHarian() berikutnya.
   Kode baru menulis 20 kolom; sheet lama punya 26. Tanpa migrasi ini, data akan
   tertulis di kolom yang salah.
   Yang dihapus: 6 kolom milik Chess.com (dicari lewat NAMA header, bukan nomor),
   dan SELURUH isi sheet 'Partai' (sheet itu memang khusus partai Chess.com).
   Seluruh spreadsheet disalin lebih dulu sebagai cadangan di Drive. */
function migrasiHapusKolomChesscom() {
  const ss = SpreadsheetApp.getActive();
  const cadangan = ss.copy('CADANGAN sebelum migrasi Chess.com - ' +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'));
  console.log('Cadangan: ' + cadangan.getUrl());

  const KOLOM_CC = ['Rush percobaan', 'Rush skor', 'Taktik tertinggi (CC)',
                    'Rekor baru?', 'Rating rapid', 'Rating blitz'];
  const shH = ss.getSheetByName(SHEET_HARIAN);
  const dihapus = [];
  if (shH) {
    // Hapus dari kanan ke kiri supaya indeks kolom lain tidak bergeser.
    KOLOM_CC.map(function (nama) {
      const header = shH.getRange(1, 1, 1, shH.getLastColumn()).getValues()[0]
        .map(function (h) { return String(h).trim(); });
      return { nama: nama, i: header.indexOf(nama) };
    }).filter(function (x) { return x.i >= 0; })
      .sort(function (a, b) { return b.i - a.i; })
      .forEach(function (x) {
        const header = shH.getRange(1, 1, 1, shH.getLastColumn()).getValues()[0]
          .map(function (h) { return String(h).trim(); });
        const i = header.indexOf(x.nama);
        if (i >= 0) { shH.deleteColumn(i + 1); dihapus.push(x.nama); }
      });
  }

  const shP = ss.getSheetByName(SHEET_PARTAI);
  let selP = 0;
  if (shP && shP.getLastRow() > 1) {
    selP = shP.getLastRow() - 1;
    shP.getRange(2, 1, selP, shP.getLastColumn()).clearContent();
  }

  const sisa = shH ? shH.getRange(1, 1, 1, shH.getLastColumn()).getValues()[0].join(' | ') : '(sheet tak ada)';
  console.log('Kolom dihapus: ' + (dihapus.join(', ') || '(tidak ada / sudah pernah dimigrasi)'));
  console.log('Partai dibersihkan: ' + selP + ' baris.');
  console.log('Header sekarang: ' + sisa);
  console.log('Kalau salah, pulihkan dari cadangan di atas.');
  log_('MIGRASI', 'Kolom Chess.com dihapus (' + dihapus.length + '). Partai: ' + selP + ' baris. Cadangan: ' + cadangan.getUrl());
}

/* Uji token Lichess dari sisi backend. Jalankan fungsi ini lalu baca Execution log.
   Menjawab pasti: token yang dipakai yang mana, dan Lichess menerimanya atau tidak. */
function tesTokenLichess() {
  var t = _tokenLi_();
  if (!t) { console.log('TIDAK ADA TOKEN. Isi Script Properties LICHESS_TOKEN, atau LICHESS_TOKEN_HARDCODE.'); return; }
  var sumber = PropertiesService.getScriptProperties().getProperty('LICHESS_TOKEN')
    ? 'Script Properties' : 'LICHESS_TOKEN_HARDCODE di kode';
  console.log('Sumber token: ' + sumber + ' | awalan: ' + t.slice(0, 8) + '...');
  var res = UrlFetchApp.fetch('https://lichess.org/api/puzzle/activity?max=5', {
    method: 'get', muteHttpExceptions: true,
    headers: { 'Accept': 'application/x-ndjson', 'Authorization': 'Bearer ' + t },
  });
  var kode = res.getResponseCode();
  console.log('HTTP ' + kode);
  if (kode === 200) {
    var n = res.getContentText().split('\n').filter(String).length;
    console.log('TOKEN VALID. ' + n + ' baris aktivitas puzzle terbaca.');
  } else if (kode === 401) {
    console.log('DITOLAK (401). Token dicabut, salah ketik, atau tanpa izin puzzle:read.');
    console.log('Buat token baru di lichess.org/account/oauth/token (centang puzzle:read SAJA).');
  } else {
    console.log('Gagal: ' + res.getContentText().slice(0, 200));
  }
}

function cabutTokenLichess() {
  PropertiesService.getScriptProperties().deleteProperty('LICHESS_TOKEN');
  console.log('Token Lichess dihapus dari Script Properties.');
}

/* Jumlah puzzle per TEMA pada satu tanggal, dari /api/puzzle/activity (ber-token).
   Skema Lichess: tiap baris NDJSON = { date(ms), win, puzzle:{ themes:[...] } }. */
function ambilTemaLichess_(tanggal, tz) {
  const out = { tema: {}, total: 0, status: 'MATI' };
  const t = _tokenLi_();
  if (!t)            { out.status = 'TANPA TOKEN'; return out; }
  if (!ANAK.lichess) { out.status = 'MATI';        return out; }
  try {
    const res = UrlFetchApp.fetch('https://lichess.org/api/puzzle/activity?max=200', {
      method: 'get', muteHttpExceptions: true,
      headers: { 'Accept': 'application/x-ndjson', 'Authorization': 'Bearer ' + t },
    });
    const kode = res.getResponseCode();
    if (kode === 401) { out.status = 'TOKEN DITOLAK';
      log_('PERINGATAN', 'Token Lichess ditolak (401) atau tanpa izin puzzle:read.'); return out; }
    if (kode !== 200) { out.status = 'HTTP ' + kode; return out; }
    res.getContentText().split('\n').filter(String).forEach(function (baris) {
      var a; try { a = JSON.parse(baris); } catch (e) { return; }
      if (!a || a.date == null || !a.puzzle) return;
      var d = new Date(msLi_(a.date));
      if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== tanggal) return;
      (a.puzzle.themes || []).forEach(function (th) { out.tema[th] = (out.tema[th] || 0) + 1; });
      out.total++;
    });
    out.status = 'OK';
  } catch (e) {
    out.status = 'GAGAL: ' + e.message;
    log_('PERINGATAN', 'Tema Lichess: ' + e.message);
  }
  return out;
}

/* Sheet lama tidak punya kolom 'Tema (JSON)'. Pastikan header kolom ke-26 ada. */
function pastikanKolomTema_(sh) {
  if (sh.getRange(1, 20).getValue() !== 'Tema (JSON)') {
    sh.getRange(1, 20).setValue('Tema (JSON)')
      .setFontWeight('bold').setBackground('#1B2130').setFontColor('#F1F2EB');
  }
}


// ═══════════════════ PENCATATAN HARIAN ═══════════════════
/** Dipanggil trigger harian. Mengulang untuk SETIAP anak di KONFIG.anak. */
function catatHarian() {
  // Cegah duplikasi baris kalau trigger overlap (eksekusi lambat / dobel).
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { log_('LEWAT', 'Eksekusi lain sedang menulis; dilewati.'); return; }
  try {
    KONFIG.anak.forEach(function (a) {
      ANAK = a;
      try { catatSatuAnak_(); }
      catch (e) { log_('GAGAL', a.nama + ': ' + e.message); }
      Utilities.sleep(800);   // sopan terhadap kedua API
    });
  } finally { lock.releaseLock(); }
}

function catatSatuAnak_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const hariIni = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  let status = 'OK';

  const li = ambilLichess_(hariIni, tz);
  if (li.status !== 'OK' && li.status !== 'MATI') status += ' | Lichess ' + li.status;

  // Verifikasi TEMA (butuh token puzzle:read di Script Properties).
  const temaLi = ambilTemaLichess_(hariIni, tz);
  li.tema = temaLi.tema; li.temaStatus = temaLi.status;
  if (['OK', 'MATI', 'TANPA TOKEN'].indexOf(temaLi.status) < 0) status += ' | Tema ' + temaLi.status;

  const baris = tulisBaris_(hariIni, status, null, [], li);
  kirimLaporan_(hariIni, baris);
}

function tulisBaris_(tanggal, status, stats, partai, li) {
  li = li || { puzzle: 0, benar: 0, salah: 0, ratingPuzzle: '', puzzleTotal: null, partai: 0, partaiList: [], tema: {}, temaStatus: '' };
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName(SHEET_HARIAN);
  pastikanKolomTema_(sh);

  // Tema hanya ditulis kalau BENAR-benar terverifikasi (status OK). Kalau tanpa
  // token / gagal, kolom dikosongkan -> frontend jatuh ke centang manual, bukan
  // mengaku 'terverifikasi nol'.
  const temaJson = (li.temaStatus === 'OK') ? JSON.stringify(li.tema || {}) : '';

  const r = ringkasPartai_(partai);

  /* BUG YANG DIPERBAIKI: partai Lichess tidak pernah masuk hitungan sama sekali.
     Anak bisa main 6 rapid di Lichess dan dasbor melaporkan "0 partai".
     Pagar anti-tilt dan batas bullet ikut buta karenanya. */
  (li.partaiList || []).forEach(g => {
    r.total++;
    if (g.speed === 'rapid')  r.rapid++;
    if (g.speed === 'blitz')  r.blitz++;
    if (g.speed === 'bullet') r.bullet++;
    if (g.res === 'M') r.menang++;
    else if (g.res === 'R') r.seri++;
    else r.kalah++;
  });
  // deret kekalahan dihitung ulang lintas situs, urut waktu
  const semua = []
    .concat(partai.map(function (g) {
      const me = (ANAK.chess || '').toLowerCase();
      const putih = (g.white.username || '').toLowerCase() === me;
      const sisi = putih ? g.white : g.black;
      const seri = ['agreed','repetition','stalemate','insufficient','50move','timevsinsufficient'];
      return { t: g.end_time * 1000,
               res: sisi.result === 'win' ? 'M' : (seri.indexOf(sisi.result) >= 0 ? 'R' : 'K') };
    }))
    .concat(li.partaiList || [])
    .sort(function (a, b) { return a.t - b.t; });
  let cur = 0; r.kalahBeruntun = 0;
  semua.forEach(function (g) {
    if (g.res === 'K') { cur++; r.kalahBeruntun = Math.max(r.kalahBeruntun, cur); }
    else cur = 0;
  });

  const T = KONFIG.target;

  // Kepatuhan: hanya menghitung yang BISA diverifikasi. Tugas manual tidak
  // ikut dihitung di sini — kalau ikut, angkanya bohong.
  const cek = [];
  if (T.puzzle > 0)        cek.push(li.puzzle >= T.puzzle);
  cek.push(r.rapid >= T.rapid);
  cek.push(r.bullet <= T.bulletMaks);
  cek.push(r.total <= T.partaiMaks);
  const kepatuhan = cek.length ? Math.round(100 * cek.filter(Boolean).length / cek.length) : 0;

  const langgar = [];
  if (T.puzzle > 0 && li.puzzle < T.puzzle) langgar.push('puzzle ' + li.puzzle + '/' + T.puzzle);
  if (r.bullet > T.bulletMaks)      langgar.push('bullet ' + r.bullet);
  if (r.total > T.partaiMaks)       langgar.push('volume ' + r.total);
  if (r.kalahBeruntun >= 3)         langgar.push('tilt ' + r.kalahBeruntun + ' kalah beruntun');

  const row = [
    tanggal, ANAK.nama, status,
    li.puzzle, li.benar, li.salah, li.ratingPuzzle,
    li.puzzleTotal == null ? '' : li.puzzleTotal,
    r.total, r.rapid, r.blitz, r.bullet, r.menang, r.kalah, r.seri,
    r.kalahBeruntun, r.akurasi || '', kepatuhan, langgar.join('; '),
    temaJson,
  ];

  // Satu baris per (tanggal, anak). Tanpa nama anak di kunci, anak kedua akan
  // MENIMPA baris anak pertama setiap hari — dan datanya hilang tanpa jejak.
  const data = sh.getDataRange().getValues();
  let idx = -1;
  for (let i = 1; i < data.length; i++)
    if (tglStr_(data[i][0], tz) === tanggal && data[i][1] === ANAK.nama) { idx = i + 1; break; }
  if (idx > 0) sh.getRange(idx, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);

  return { tanggal, li, kepatuhan, langgar, r, status };
}

function ringkasPartai_(partai) {
  const me = (ANAK.chess || '').toLowerCase();
  const out = { total: partai.length, rapid: 0, blitz: 0, bullet: 0,
                menang: 0, kalah: 0, seri: 0, kalahBeruntun: 0, akurasi: null };
  const seri = ['agreed', 'repetition', 'stalemate', 'insufficient', '50move', 'timevsinsufficient'];
  const akur = [];
  let beruntun = 0;

  partai.slice().sort((a, b) => a.end_time - b.end_time).forEach(g => {
    if (g.time_class === 'rapid')  out.rapid++;
    if (g.time_class === 'blitz')  out.blitz++;
    if (g.time_class === 'bullet') out.bullet++;

    const putih = (g.white.username || '').toLowerCase() === me;
    const sisi = putih ? g.white : g.black;

    if (sisi.result === 'win') { out.menang++; beruntun = 0; }
    else if (seri.indexOf(sisi.result) >= 0) { out.seri++; beruntun = 0; }
    else { out.kalah++; beruntun++; out.kalahBeruntun = Math.max(out.kalahBeruntun, beruntun); }

    if (g.accuracies) {
      const a = putih ? g.accuracies.white : g.accuracies.black;
      if (a) akur.push(a);
    }
  });
  if (akur.length) out.akurasi = Math.round(10 * akur.reduce((x, y) => x + y, 0) / akur.length) / 10;
  return out;
}


// ═══════════════════ LAPORAN ═══════════════════
function kirimLaporan_(tanggal, b) {
  const T = KONFIG.target;
  const lulus = x => x ? '✓' : '✗';

  const badan = [
    'LAPORAN HARIAN — ' + ANAK.nama + ' — ' + tanggal,
    '',
    'Kepatuhan terverifikasi: ' + b.kepatuhan + '%',
    '',
    T.puzzle > 0
      ? lulus(b.li.puzzle >= T.puzzle) + ' Puzzle (Lichess): ' + b.li.puzzle + '/' + T.puzzle +
        '  (' + b.li.benar + ' benar, ' + b.li.salah + ' salah' +
        (b.li.ratingPuzzle ? ', rating ' + b.li.ratingPuzzle : '') + ')'
      : '',
    lulus(b.r.rapid >= T.rapid) + ' Partai rapid: ' + b.r.rapid + '/' + T.rapid,
    lulus(b.r.bullet <= T.bulletMaks) + ' Bullet: ' + b.r.bullet + ' (batas ' + T.bulletMaks + ')',
    lulus(b.r.total <= T.partaiMaks) + ' Total partai: ' + b.r.total + ' (batas ' + T.partaiMaks + ')',
    '',
    'Hasil hari ini: ' + b.r.menang + 'M / ' + b.r.seri + 'R / ' + b.r.kalah + 'K' +
      (b.r.akurasi ? '  ·  akurasi rata-rata ' + b.r.akurasi + '%' : ''),
    '',
    b.langgar.length ? 'PERLU PERHATIAN:\n- ' + b.langgar.join('\n- ') : 'Tidak ada pelanggaran.',
    '',
    b.r.kalahBeruntun >= 3
      ? 'CATATAN: ' + b.r.kalahBeruntun + ' kekalahan beruntun. Bermain terus setelah tiga kekalahan ' +
        'hampir selalu memperburuk rating dan suasana hati. Hentikan sesi, jangan tambah partai.'
      : '',
    '',
    '--',
    'Puzzle dan partai dihitung dari Lichess.',
    'Pelajaran, video, dan latihan papan fisik TIDAK terlihat di sini sama sekali.',
    'Untuk itu, tanya anaknya. Dasbor tidak menggantikan percakapan.',
  ].filter(x => x !== '').join('\n');

  kirimNotif_('[Road To Grand Master] ' + ANAK.nama + ' — ' + tanggal + ' — kepatuhan ' + b.kepatuhan + '%', badan);
}

/* Satu pintu keluar untuk semua pemberitahuan. Email + Telegram. */
function kirimNotif_(subjek, badan) {
  kirimEmail_(subjek, badan);
  kirimTelegram_(subjek + '\n\n' + badan);
}

function kirimTelegram_(pesan) {
  if (!KONFIG.telegramToken || !KONFIG.telegramChatId) return;
  try {
    const res = UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + KONFIG.telegramToken + '/sendMessage', {
        method: 'post',
        muteHttpExceptions: true,
        payload: {
          chat_id: KONFIG.telegramChatId,
          text: pesan.slice(0, 4000),   // batas Telegram 4096 karakter
          disable_web_page_preview: 'true',
        },
      });
    const kode = res.getResponseCode();
    if (kode !== 200) log_('PERINGATAN', 'Telegram HTTP ' + kode + ': ' + res.getContentText().slice(0, 200));
  } catch (e) {
    log_('PERINGATAN', 'Telegram: ' + e.message);
  }
}

function kirimEmail_(subjek, badan) {
  const to = [].concat(KONFIG.emailOrangTua, KONFIG.emailPelatih)
    .filter(e => e && e.indexOf('@') > 0 && e.indexOf('ganti@') !== 0);
  if (!to.length) return;
  try { MailApp.sendEmail(to.join(','), subjek, badan); }
  catch (e) { log_('GAGAL', 'Email: ' + e.message); }
}

function log_(tingkat, pesan) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  if (sh) sh.appendRow([new Date(), tingkat, pesan]);
  console.log(tingkat + ': ' + pesan);
}

// ═══════════════════ ENDPOINT UNTUK DASBOR HTML ═══════════════════
/**
 * Deploy > New deployment > Web app
 *   Execute as       : Me
 *   Who has access   : Anyone
 * Salin URL /exec, tempel ke kolom "URL Web App" di dasbor HTML.
 *
 * Dua masalah selesai sekaligus:
 *   1. CORS — permintaan ke Google, bukan langsung ke API catur.
 *   2. Riwayat — dasbor mendapat SELURUH catatan harian, termasuk hari-hari
 *      saat dasbor tidak dibuka sama sekali.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.ai) {
    return ContentService
      .createTextOutput(JSON.stringify({ ai: coachAIGemini(e.parameter.d || "", e.parameter.j || "review") }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const minta = (e && e.parameter && e.parameter.anak) || null;
  ANAK = KONFIG.anak.filter(function (a) {
    return !minta || a.nama.toLowerCase() === String(minta).toLowerCase();
  })[0] || KONFIG.anak[0];

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_HARIAN);
  const tzG = ss.getSpreadsheetTimeZone();
  const data = sh ? sh.getDataRange().getValues() : [];
  const snaps = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[1] !== ANAK.nama) continue;   // hanya anak yang diminta
    const kunci = tglStr_(r[0], tzG);   // Date -> "yyyy-MM-dd", apa pun bentuk aslinya
    if (!/^\d{4}-\d{2}-\d{2}$/.test(kunci)) continue;
    snaps[kunci] = {
      puzzles: r[3] || 0, puzzleWin: r[4] || 0, puzzleLoss: r[5] || 0, puzzleRating: r[6] || null,
      puzTotal: r[7] === '' || r[7] == null ? null : Number(r[7]),
      status: r[2], kepatuhan: r[17], langgar: r[18],
      tema: (function () { var v = r[19]; if (v == null || String(v).trim() === '') return null;
        try { return JSON.parse(v); } catch (e) { return null; } })(),
    };
  }
  let stats = null, games = [], lichess = null;
  if (ANAK.lichess) {
    try {
      lichess = {
        user: ANAK.lichess,
        prof: ambil_(LI()),
        act:  ambil_(LI() + '/activity'),
        hist: ambil_(LI() + '/rating-history'),
      };
    } catch (e) { /* abaikan */ }
  }

  // Tema live hari ini & beberapa hari terakhir — server-side, tanpa token di browser.
  var temaStatus = 'TIDAK DIJALANKAN';
  try {
    var tokLi = _tokenLi_();
    temaStatus = tokLi ? 'MENCOBA' : 'TANPA TOKEN';
    if (tokLi && ANAK.lichess) {
      var resT = UrlFetchApp.fetch('https://lichess.org/api/puzzle/activity?max=200', {
        method: 'get', muteHttpExceptions: true,
        headers: { 'Accept': 'application/x-ndjson', 'Authorization': 'Bearer ' + tokLi },
      });
      var kodeT = resT.getResponseCode();
      temaStatus = kodeT === 200 ? 'OK' : (kodeT === 401 ? 'TOKEN DITOLAK (401)' : 'HTTP ' + kodeT);
      if (kodeT === 200) {
        var perHari = {};
        resT.getContentText().split('\n').filter(String).forEach(function (baris) {
          var a; try { a = JSON.parse(baris); } catch (e) { return; }
          if (!a || a.date == null || !a.puzzle) return;
          var k = Utilities.formatDate(new Date(msLi_(a.date)), tzG, 'yyyy-MM-dd');
          perHari[k] = perHari[k] || {};
          (a.puzzle.themes || []).forEach(function (th) { perHari[k][th] = (perHari[k][th] || 0) + 1; });
        });
        Object.keys(perHari).forEach(function (k) {
          snaps[k] = snaps[k] || {};
          snaps[k].tema = perHari[k];   // live menimpa yang tersimpan (lebih baru)
        });
        // Tandai HARI INI "verifikasi aktif" walau 0 puzzle, agar tugas tema
        // tampil OTOMATIS (0/N) di dasbor, bukan MANUAL.
        var todayG = Utilities.formatDate(new Date(), tzG, 'yyyy-MM-dd');
        snaps[todayG] = snaps[todayG] || {};
        if (snaps[todayG].tema == null) snaps[todayG].tema = {};
      }
    }
  } catch (e) { /* abaikan; tema tersimpan dari Sheet tetap dipakai */ }

  return ContentService
    .createTextOutput(JSON.stringify({
      temaStatus: temaStatus,
      nama: ANAK.nama,
      daftarAnak: KONFIG.anak.map(function (a) { return a.nama; }),
      stats: stats, games: games, lichess: lichess, snaps: snaps,
      diambil: new Date().toISOString(),
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════ UJI MANUAL ═══════════════════
/** Uji Telegram sendiri. Jalankan setelah mengisi token & chat id. */
function ujiTelegram() {
  if (!KONFIG.telegramToken || !KONFIG.telegramChatId) {
    const m = 'Telegram tidak diaktifkan (token atau chatId kosong).';
    console.log(m); return m;
  }
  kirimTelegram_('Uji koneksi Road To Grand Master — kalau pesan ini sampai, laporan malam akan masuk ke sini.');
  const m = 'Pesan uji dikirim. Cek Telegram.';
  console.log(m); return m;
}

/** Jalankan ini dulu sebelum pasang(). Menguji KEDUA API. */
function ujiKoneksi() {
  let semua = '';
  KONFIG.anak.forEach(function (a) {
    ANAK = a;
    semua += '\n===== ' + a.nama + ' =====\n' + ujiSatu_() + '\n';
  });
  console.log(semua);
  return semua;
}
function ujiSatu_() {
  let pesan = '';
  // ── Lichess (sumber puzzle — yang paling penting) ──
  try {
    const p = ambil_(LI());
    const a = ambil_(LI() + '/activity');
    pesan += 'LICHESS BERHASIL.\n' +
      '  Rating puzzle: ' + (p.perfs && p.perfs.puzzle ? p.perfs.puzzle.rating : '-') + '\n' +
      '  Total puzzle : ' + (p.perfs && p.perfs.puzzle ? p.perfs.puzzle.games : 0) + '\n' +
      '  Hari aktif   : ' + ((a && a.length) || 0) + '\n\n';
  } catch (e) {
    pesan += 'LICHESS GAGAL: ' + e.message + '\n\n';
  }
  return pesan;
}

/* ============ PELATIH AI (Gemini, gratis) ============
   Kunci disimpan di Properti Skrip (GEMINI_KEY), TIDAK di file publik.
   Frontend memanggil: <URL>/exec?ai=1&d=<json angka performa tanpa nama>.
   Model bisa diganti via properti GEMINI_MODEL (default gemini-2.5-flash). */
function coachAIGemini(dataStr, jenis) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty("GEMINI_KEY");
  if (!key) return "ERROR: GEMINI_KEY belum diatur di Properti Skrip.";
  var model = props.getProperty("GEMINI_MODEL") || "gemini-2.5-flash-lite";
  var prompt;
  if (jenis === "kurikulum") {
    prompt =
      "Kamu pelatih catur untuk seorang ANAK. Data di bawah berisi ringkasan statistik DAN beberapa partai rapid terbaru " +
      "lengkap dengan data mesin (akurasi, kesalahan serius per fase) dan notasi langkah (moves). " +
      "UTAMAKAN DATA MESIN. Notasi hanya untuk konteks pola pembukaan; JANGAN memvonis kualitas langkah spesifik sebagai benar/salah, " +
      "karena analisis caturmu bisa keliru — cukup kenali pola berulang (mis. sering kalah cepat dengan pembukaan tertentu, banjir kesalahan di fase tertentu). " +
      "Tugas: susun RENCANA LATIHAN UNTUK HARI INI, diturunkan dari partai-partai terbaru itu. " +
      "Beri 3-5 tugas konkret hari ini: tema puzzle, review pembukaan tertentu bila terlihat pola, aturan tempo bila banyak langkah tanpa pikir, plus jumlah puzzle. " +
      "Ringkas, daftar bernomor, Bahasa Indonesia. Akhiri satu kalimat penyemangat untuk anak.\n\nDATA:\n" + dataStr;
  } else {
    prompt =
      "Kamu pelatih catur untuk seorang ANAK. Data di bawah berisi ringkasan statistik DAN beberapa partai terbaru " +
      "dengan data mesin dan notasi langkah. UTAMAKAN DATA MESIN (akurasi, kesalahan per fase). " +
      "Notasi hanya konteks; jangan memvonis langkah spesifik benar/salah karena analisis caturmu bisa keliru. Tanpa menyebut nama. " +
      "Jawab dalam Bahasa Indonesia yang hangat dan ringkas, dalam 3 bagian pendek:\n" +
      "1) Diagnosis: satu-dua kalimat, apa yang paling menahan hasil.\n" +
      "2) Latihan minggu ini: dua-tiga saran konkret dan bisa dikerjakan.\n" +
      "3) Untuk anak: satu kalimat penyemangat.\n" +
      "Ini pendapat berbasis data, hindari klaim pasti.\n\nDATA:\n" + dataStr;
  }
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent";
  var opt = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": key },
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  };
  var res = UrlFetchApp.fetch(url, opt);
  var code = res.getResponseCode();
  if (code === 429) { Utilities.sleep(3000); res = UrlFetchApp.fetch(url, opt); code = res.getResponseCode(); }
  var body = res.getContentText();
  if (code !== 200) return "ERROR " + code + ": " + body.slice(0, 500);
  try {
    var j = JSON.parse(body);
    var t = j.candidates && j.candidates[0] && j.candidates[0].content &&
            j.candidates[0].content.parts && j.candidates[0].content.parts[0] &&
            j.candidates[0].content.parts[0].text;
    return t || ("ERROR: respons kosong. " + body.slice(0, 300));
  } catch (err) {
    return "ERROR parse: " + err;
  }
}

// Jalankan ini dari editor (pilih tesGeminiCall lalu Run) untuk menguji kunci & model.
function tesGeminiCall() {
  var out = coachAIGemini('{"partai_rapid":20,"menang_persen":55,"akurasi_rata":78,"fase_terlemah":"pembukaan","salah_dimainkan_tanpa_pikir":6}');
  Logger.log(out);
}
