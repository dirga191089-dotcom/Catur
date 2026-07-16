/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ROAD TO GRAND MASTER · PANTAU — Pencatat Harian (Google Apps Script)
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * PERAN SKRIP INI — dan yang BUKAN perannya.
 *
 * Perannya ada satu: menjadi SAKSI yang tidak pernah lupa.
 * Chess.com tidak menyimpan riwayat Puzzle Rush maupun riwayat rating taktik.
 * Ia hanya memberi keadaan SAAT INI. Kalau tidak ada yang memotretnya setiap
 * hari, hari itu hilang selamanya. Browser tidak bisa memotret saat tidak
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
 * Chess.com ada di belakang Cloudflare. Permintaan dari IP Google Apps Script
 * KADANG diblokir dengan HTTP 403. Kalau itu terjadi, skrip ini TIDAK akan
 * menulis angka nol seolah anak tidak latihan — ia menulis "GAGAL" di kolom
 * status dan mengirim email peringatan. Kesalahan diam adalah kesalahan
 * terburuk untuk aplikasi pengawasan.
 */

// ═══════════════════ KONFIG ═══════════════════
const KONFIG = {
  /* Banyak anak. Tambahkan objek baru ke daftar ini; tidak perlu apa pun lagi.
     Setiap anak mendapat barisnya sendiri di sheet, dibedakan kolom "Anak". */
  anak: [
    { nama: 'Varisha', chess: 'VarishaArbas', lichess: 'VarishaChess' },
    // { nama: 'Adik',  chess: '',             lichess: '' },
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

  // Kontak untuk header User-Agent. Chess.com MEMINTA ini, dan mengisinya
  // memperkecil kemungkinan diblokir. Jangan dikosongkan.
  kontak: 'ganti@email-anda.com',

  // Target harian. Harus sama dengan yang di dasbor HTML.
  target: {
    puzzle:         30,  // puzzle per hari di Lichess — SATU-SATUNYA hitungan puzzle
    rapid:           2,  // partai rapid per hari         yang bisa diverifikasi API
    bulletMaks:      0,  // batas partai bullet (pagar, bukan target)
    partaiMaks:      6,  // batas total partai per hari (anti-tilt)
    rushPercobaan:   0,  // Puzzle Rush Chess.com — 0 = tidak dipakai
    rushSkor:        0,
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
  backfillPartai();                     // tarik seluruh arsip partai
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
    'Rush percobaan', 'Rush skor',
    'Taktik tertinggi (CC)', 'Rekor baru?', 'Rating rapid', 'Rating blitz',
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
      // Chess.com meminta User-Agent berisi kontak. Ini bukan basa-basi:
      // permintaan tanpa ini lebih sering diblokir Cloudflare.
      'User-Agent': 'RTGM-Pantau/1.0 (kontak: ' + KONFIG.kontak + ')',
      'Accept': 'application/json',
    },
  });
  const kode = res.getResponseCode();
  if (kode === 403) {
    throw new Error('403 — Cloudflare Chess.com memblokir IP Google Apps Script. ' +
      'Ini terjadi berkala dan bukan kesalahan username. Coba lagi besok; kalau menetap, ' +
      'gunakan dasbor HTML langsung dari browser (browser tidak diblokir).');
  }
  if (kode === 404) throw new Error('404 — username tidak ditemukan: ' + ANAK.chess);
  if (kode === 429) throw new Error('429 — terlalu banyak permintaan. Kurangi frekuensi trigger.');
  if (kode !== 200) throw new Error('HTTP ' + kode + ' dari ' + url);
  return JSON.parse(res.getContentText());
}

/* Anak yang sedang diproses. Diset oleh catatHarian() sebelum tiap putaran. */
let ANAK = KONFIG.anak[0];
const BASE = () => 'https://api.chess.com/pub/player/' + (ANAK.chess || '').toLowerCase();
const LI   = () => 'https://lichess.org/api/user/' + (ANAK.lichess || '');

/**
 * Lichess: SATU-SATUNYA sumber jumlah puzzle harian yang bisa diverifikasi.
 * PENTING: timestamp Lichess dalam MILIDETIK, Chess.com dalam DETIK.
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
   ini harus milik akun Lichess si anak (Varisha). Banyak anak = butuh mekanisme
   token per anak; untuk satu anak, ini sudah benar. */
function _tokenLi_() {
  return PropertiesService.getScriptProperties().getProperty('LICHESS_TOKEN') || '';
}
function pasangTokenLichess() {
  const t = '';                 // <- tempel token puzzle:read di sini, jalankan SEKALI
  if (!t) { console.log('Isi variabel t di dalam fungsi ini dulu, lalu jalankan.'); return; }
  PropertiesService.getScriptProperties().setProperty('LICHESS_TOKEN', t.trim());
  console.log('Token tersimpan di Script Properties. SEKARANG kosongkan lagi variabel t di kode ini.');
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
  if (sh.getRange(1, 26).getValue() !== 'Tema (JSON)') {
    sh.getRange(1, 26).setValue('Tema (JSON)')
      .setFontWeight('bold').setBackground('#1B2130').setFontColor('#F1F2EB');
  }
}

function ambilStats_()    { return ambil_(BASE() + '/stats'); }
function ambilArsip_()    { return ambil_(BASE() + '/games/archives').archives || []; }
function ambilBulan_(url) { return ambil_(url).games || []; }

// ═══════════════════ PENCATATAN HARIAN ═══════════════════
/** Dipanggil trigger harian. Mengulang untuk SETIAP anak di KONFIG.anak. */
function catatHarian() {
  KONFIG.anak.forEach(function (a) {
    ANAK = a;
    try { catatSatuAnak_(); }
    catch (e) { log_('GAGAL', a.nama + ': ' + e.message); }
    Utilities.sleep(800);   // sopan terhadap kedua API
  });
}

function catatSatuAnak_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const hariIni = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  let stats, partaiHariIni = [], status = 'OK';
  try {
    stats = ambilStats_();
  } catch (e) {
    // Chess.com gagal TIDAK boleh membuang data Lichess. Puzzle tetap dicatat.
    log_('GAGAL', 'Ambil stats Chess.com: ' + e.message);
    const liSaja = ambilLichess_(hariIni, tz);
    const temaSaja = ambilTemaLichess_(hariIni, tz);
    liSaja.tema = temaSaja.tema; liSaja.temaStatus = temaSaja.status;
    tulisBaris_(hariIni, 'GAGAL Chess.com: ' + e.message, null, [], liSaja);
    kirimNotif_('[Road To Grand Master] Chess.com GAGAL — ' + ANAK.nama + ' ' + hariIni,
      'Data Chess.com tidak bisa ditarik hari ini.\n\n' + e.message +
      '\n\nData Lichess TETAP tercatat: ' + liSaja.puzzle + ' puzzle.' +
      '\n\nAngka partai untuk ' + hariIni + ' KOSONG karena pengambilannya gagal — ' +
      'bukan karena anak tidak bermain. Jangan menghukum atas kesalahan kode.');
    return;
  }

  // Partai hari ini diambil dari arsip bulan berjalan.
  try {
    const arsip = ambilArsip_();
    if (arsip.length) {
      const semua = ambilBulan_(arsip[arsip.length - 1]);
      partaiHariIni = semua.filter(g =>
        Utilities.formatDate(new Date(g.end_time * 1000), tz, 'yyyy-MM-dd') === hariIni);
      simpanPartai_(semua, tz);
    }
  } catch (e) {
    status = 'SEBAGIAN: arsip partai gagal (' + e.message + ')';
    log_('PERINGATAN', 'Arsip gagal: ' + e.message);
  }

  const li = ambilLichess_(hariIni, tz);
  if (li.status !== 'OK' && li.status !== 'MATI') status += ' | Lichess ' + li.status;

  // Verifikasi TEMA (butuh token puzzle:read di Script Properties).
  const temaLi = ambilTemaLichess_(hariIni, tz);
  li.tema = temaLi.tema; li.temaStatus = temaLi.status;
  if (['OK', 'MATI', 'TANPA TOKEN'].indexOf(temaLi.status) < 0) status += ' | Tema ' + temaLi.status;

  const baris = tulisBaris_(hariIni, status, stats, partaiHariIni, li);
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

  const rush  = stats && stats.puzzle_rush && stats.puzzle_rush.daily ? stats.puzzle_rush.daily : null;
  const rBest = stats && stats.puzzle_rush && stats.puzzle_rush.best  ? stats.puzzle_rush.best  : null;
  const tac   = stats && stats.tactics && stats.tactics.highest ? stats.tactics.highest : null;

  const rushPercobaan = rush ? (rush.total_attempts || 0) : 0;
  const rushSkor      = rush ? (rush.score || 0) : 0;
  const taktikTinggi  = tac ? tac.rating : '';
  const rekorBaru     = tac && Utilities.formatDate(new Date(tac.date * 1000), tz, 'yyyy-MM-dd') === tanggal;

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
      const me = ANAK.chess.toLowerCase();
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
  if (T.rushPercobaan > 0) cek.push(rushPercobaan >= T.rushPercobaan);
  cek.push(r.rapid >= T.rapid);
  cek.push(r.bullet <= T.bulletMaks);
  cek.push(r.total <= T.partaiMaks);
  const kepatuhan = cek.length ? Math.round(100 * cek.filter(Boolean).length / cek.length) : 0;

  const langgar = [];
  if (T.puzzle > 0 && li.puzzle < T.puzzle) langgar.push('puzzle ' + li.puzzle + '/' + T.puzzle);
  if (r.bullet > T.bulletMaks)      langgar.push('bullet ' + r.bullet);
  if (r.total > T.partaiMaks)       langgar.push('volume ' + r.total);
  if (r.kalahBeruntun >= 3)         langgar.push('tilt ' + r.kalahBeruntun + ' kalah beruntun');
  if (rushPercobaan === 0)          langgar.push('Rush tidak dikerjakan');
  if (rushPercobaan > 0 && rushSkor < T.rushSkor) langgar.push('skor Rush ' + rushSkor + ' < ' + T.rushSkor);

  const row = [
    tanggal, ANAK.nama, status,
    li.puzzle, li.benar, li.salah, li.ratingPuzzle,
    li.puzzleTotal == null ? '' : li.puzzleTotal,
    rushPercobaan, rushSkor,
    taktikTinggi, rekorBaru ? 'YA' : '',
    stats && stats.chess_rapid && stats.chess_rapid.last ? stats.chess_rapid.last.rating : '',
    stats && stats.chess_blitz && stats.chess_blitz.last ? stats.chess_blitz.last.rating : '',
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

  return { tanggal, li, rushPercobaan, rushSkor, rekorBaru, kepatuhan, langgar, r, status };
}

function ringkasPartai_(partai) {
  const me = ANAK.chess.toLowerCase();
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

function simpanPartai_(partai, tz) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PARTAI);
  const me = ANAK.chess.toLowerCase();
  const adaUrl = {};
  sh.getDataRange().getValues().slice(1).forEach(r => adaUrl[r[9]] = true);

  const seri = ['agreed', 'repetition', 'stalemate', 'insufficient', '50move', 'timevsinsufficient'];
  const baru = partai.filter(g => !adaUrl[g.url]).sort((a, b) => a.end_time - b.end_time).map(g => {
    const putih = (g.white.username || '').toLowerCase() === me;
    const sisi = putih ? g.white : g.black;
    const lawan = putih ? g.black : g.white;
    const hasil = sisi.result === 'win' ? 'M' : (seri.indexOf(sisi.result) >= 0 ? 'R' : 'K');
    const acc = g.accuracies ? (putih ? g.accuracies.white : g.accuracies.black) : '';
    return [
      ANAK.nama,
      Utilities.formatDate(new Date(g.end_time * 1000), tz, 'yyyy-MM-dd HH:mm'),
      g.time_class, lawan.username, lawan.rating, hasil, sisi.rating, acc || '',
      ecoDari_(g.pgn), g.url,
    ];
  });
  if (baru.length) sh.getRange(sh.getLastRow() + 1, 1, baru.length, baru[0].length).setValues(baru);
}

function ecoDari_(pgn) {
  if (!pgn) return '';
  const m = pgn.match(/\[ECO "([^"]+)"\]/);
  return m ? m[1] : '';
}

// ═══════════════════ BACKFILL ═══════════════════
/** Tarik SELURUH riwayat partai. Jalankan sekali saat pertama pasang. */
function backfillPartai() {
  KONFIG.anak.forEach(function (a) { ANAK = a; backfillSatu_(); });
}
function backfillSatu_() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  if (!ANAK.chess) return;
  let arsip;
  try { arsip = ambilArsip_(); }
  catch (e) { log_('GAGAL', 'Backfill: ' + e.message); return; }

  let n = 0;
  arsip.forEach(url => {
    try {
      const g = ambilBulan_(url);
      simpanPartai_(g, tz);
      n += g.length;
      Utilities.sleep(400);   // sopan terhadap rate limit
    } catch (e) { log_('PERINGATAN', 'Bulan gagal: ' + url + ' — ' + e.message); }
  });
  log_('OK', 'Backfill ' + ANAK.nama + ': ' + n + ' partai dari ' + arsip.length + ' bulan.');
}

// ═══════════════════ LAPORAN ═══════════════════
function kirimLaporan_(tanggal, b) {
  const T = KONFIG.target;
  const lulus = x => x ? '\u2713' : '\u2717';

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
    T.rushPercobaan > 0
      ? lulus(b.rushPercobaan >= T.rushPercobaan) + ' Puzzle Rush: ' + b.rushPercobaan + '/' + T.rushPercobaan +
        ' (skor ' + b.rushSkor + ', target ' + T.rushSkor + ')'
      : '',
    lulus(b.r.rapid >= T.rapid) + ' Partai rapid: ' + b.r.rapid + '/' + T.rapid,
    lulus(b.r.bullet <= T.bulletMaks) + ' Bullet: ' + b.r.bullet + ' (batas ' + T.bulletMaks + ')',
    lulus(b.r.total <= T.partaiMaks) + ' Total partai: ' + b.r.total + ' (batas ' + T.partaiMaks + ')',
    '',
    'Hasil hari ini: ' + b.r.menang + 'M / ' + b.r.seri + 'R / ' + b.r.kalah + 'K' +
      (b.r.akurasi ? '  ·  akurasi rata-rata ' + b.r.akurasi + '%' : ''),
    b.rekorBaru ? 'REKOR RATING TAKTIK BARU HARI INI.' : '',
    '',
    b.langgar.length ? 'PERLU PERHATIAN:\n- ' + b.langgar.join('\n- ') : 'Tidak ada pelanggaran.',
    '',
    b.r.kalahBeruntun >= 3
      ? 'CATATAN: ' + b.r.kalahBeruntun + ' kekalahan beruntun. Bermain terus setelah tiga kekalahan ' +
        'hampir selalu memperburuk rating dan suasana hati. Hentikan sesi, jangan tambah partai.'
      : '',
    '',
    '--',
    'Puzzle dihitung dari Lichess. Partai dihitung dari Chess.com + Lichess.',
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
 *   1. CORS — permintaan tidak lagi ke chess.com, tapi ke Google.
 *   2. Riwayat — dasbor mendapat SELURUH catatan harian, termasuk hari-hari
 *      saat dasbor tidak dibuka sama sekali.
 */
function doGet(e) {
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
      rushAttempts: r[8] || 0, rushScore: r[9] || 0,
      tacticsHigh: r[10] || null, rapid: r[12] || null, blitz: r[13] || null,
      status: r[2], kepatuhan: r[23], langgar: r[24],
      tema: (function () { var v = r[25]; if (v == null || String(v).trim() === '') return null;
        try { return JSON.parse(v); } catch (e) { return null; } })(),
    };
  }
  let stats = null, games = [], lichess = null;
  try { stats = ambilStats_(); } catch (e) { /* pakai catatan saja */ }
  try {
    const arsip = ambilArsip_();
    if (arsip.length) games = ambilBulan_(arsip[arsip.length - 1]);
  } catch (e) { /* abaikan */ }
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
  try {
    var tokLi = _tokenLi_();
    if (tokLi && ANAK.lichess) {
      var resT = UrlFetchApp.fetch('https://lichess.org/api/puzzle/activity?max=200', {
        method: 'get', muteHttpExceptions: true,
        headers: { 'Accept': 'application/x-ndjson', 'Authorization': 'Bearer ' + tokLi },
      });
      if (resT.getResponseCode() === 200) {
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
      }
    }
  } catch (e) { /* abaikan; tema tersimpan dari Sheet tetap dipakai */ }

  return ContentService
    .createTextOutput(JSON.stringify({
      user: (ANAK.chess || '').toLowerCase(),
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
  kirimTelegram_('Uji koneksi Road To Grand Master \u2014 kalau pesan ini sampai, laporan malam akan masuk ke sini.');
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
  // ── Chess.com (sumber partai) ──
  try {
    const s = ambilStats_();
    const rush = s.puzzle_rush && s.puzzle_rush.daily ? s.puzzle_rush.daily : null;
    pesan += 'CHESS.COM BERHASIL.\n' +
      '  Rapid           : ' + (s.chess_rapid && s.chess_rapid.last ? s.chess_rapid.last.rating : '-') + '\n' +
      '  Taktik tertinggi: ' + (s.tactics && s.tactics.highest ? s.tactics.highest.rating : '-') + '\n' +
      '  Puzzle Rush kini: ' + (rush ? rush.total_attempts + ' percobaan, skor ' + rush.score : 'belum ada');
  } catch (e) {
    pesan += 'CHESS.COM GAGAL: ' + e.message + '\n' +
      '  Kalau ini 403, Cloudflare memblokir IP Google. Puzzle tetap tercatat lewat Lichess,\n' +
      '  tapi partai tidak. Pakai dasbor browser untuk data partai.';
  }
  return pesan;
}
