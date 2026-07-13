/**
 * ═══════════════════════════════════════════════════════════════════════════
 * RTGM · PANTAU — Pencatat Harian (Google Apps Script)
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
  chessUsername: 'VarishaArbas',   // username Chess.com anak
  namaAnak:      'Varisha',

  // Email penerima laporan harian. Kosongkan array untuk mematikan email.
  emailOrangTua: ['ganti@email-anda.com'],
  emailPelatih:  [],

  // Kontak untuk header User-Agent. Chess.com MEMINTA ini, dan mengisinya
  // memperkecil kemungkinan diblokir. Jangan dikosongkan.
  kontak: 'ganti@email-anda.com',

  // Target harian. Harus sama dengan yang di dasbor HTML.
  target: {
    rushPercobaan: 20,   // minimal puzzle dicoba di Puzzle Rush
    rushSkor:      12,   // target skor Puzzle Rush
    rapid:          2,   // partai rapid per hari
    bulletMaks:     0,   // batas partai bullet (pagar, bukan target)
    partaiMaks:     6,   // batas total partai per hari (anti-tilt)
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
    'Terpasang. Skrip akan berjalan tiap hari jam ' + KONFIG.jamCatat + ':00.', 'RTGM Pantau', 8);
}

function hapusTriggerLama_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['catatHarian'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
}

function siapkanSheet_() {
  const ss = SpreadsheetApp.getActive();
  mk_(ss, SHEET_HARIAN, [
    'Tanggal', 'Status ambil', 'Rush percobaan', 'Rush skor', 'Rush terbaik',
    'Taktik tertinggi', 'Rekor baru?', 'Rating rapid', 'Rating blitz',
    'Partai hari ini', 'Rapid', 'Blitz', 'Bullet', 'Menang', 'Kalah', 'Seri',
    'Kalah beruntun', 'Akurasi rata2', 'Kepatuhan %', 'Pelanggaran',
  ]);
  mk_(ss, SHEET_PARTAI, [
    'Waktu selesai', 'Kontrol', 'Lawan', 'Rating lawan', 'Hasil',
    'Rating anak', 'Akurasi anak', 'ECO', 'URL',
  ]);
  mk_(ss, SHEET_LOG, ['Waktu', 'Tingkat', 'Pesan']);
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
  if (kode === 404) throw new Error('404 — username tidak ditemukan: ' + KONFIG.chessUsername);
  if (kode === 429) throw new Error('429 — terlalu banyak permintaan. Kurangi frekuensi trigger.');
  if (kode !== 200) throw new Error('HTTP ' + kode + ' dari ' + url);
  return JSON.parse(res.getContentText());
}

const BASE = () => 'https://api.chess.com/pub/player/' + KONFIG.chessUsername.toLowerCase();

function ambilStats_()    { return ambil_(BASE() + '/stats'); }
function ambilArsip_()    { return ambil_(BASE() + '/games/archives').archives || []; }
function ambilBulan_(url) { return ambil_(url).games || []; }

// ═══════════════════ PENCATATAN HARIAN ═══════════════════
function catatHarian() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const hariIni = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  let stats, partaiHariIni = [], status = 'OK';
  try {
    stats = ambilStats_();
  } catch (e) {
    log_('GAGAL', 'Ambil stats: ' + e.message);
    tulisBaris_(hariIni, 'GAGAL: ' + e.message, null, []);
    kirimEmail_('[RTGM] GAGAL menarik data ' + KONFIG.namaAnak,
      'Skrip tidak bisa menarik data hari ini.\n\n' + e.message +
      '\n\nTIDAK ADA data untuk ' + hariIni + '. Jangan anggap anak tidak berlatih — ' +
      'yang gagal adalah pengambilan datanya, bukan latihannya.');
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

  const baris = tulisBaris_(hariIni, status, stats, partaiHariIni);
  kirimLaporan_(hariIni, baris);
}

function tulisBaris_(tanggal, status, stats, partai) {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName(SHEET_HARIAN);

  const rush  = stats && stats.puzzle_rush && stats.puzzle_rush.daily ? stats.puzzle_rush.daily : null;
  const rBest = stats && stats.puzzle_rush && stats.puzzle_rush.best  ? stats.puzzle_rush.best  : null;
  const tac   = stats && stats.tactics && stats.tactics.highest ? stats.tactics.highest : null;

  const rushPercobaan = rush ? (rush.total_attempts || 0) : 0;
  const rushSkor      = rush ? (rush.score || 0) : 0;
  const taktikTinggi  = tac ? tac.rating : '';
  const rekorBaru     = tac && Utilities.formatDate(new Date(tac.date * 1000), tz, 'yyyy-MM-dd') === tanggal;

  const r = ringkasPartai_(partai);
  const T = KONFIG.target;

  // Kepatuhan: hanya menghitung yang BISA diverifikasi. Tugas manual tidak
  // ikut dihitung di sini — kalau ikut, angkanya bohong.
  const cek = [
    rushPercobaan >= T.rushPercobaan,
    r.rapid >= T.rapid,
    r.bullet <= T.bulletMaks,
    r.total <= T.partaiMaks,
  ];
  const kepatuhan = Math.round(100 * cek.filter(Boolean).length / cek.length);

  const langgar = [];
  if (r.bullet > T.bulletMaks)      langgar.push('bullet ' + r.bullet);
  if (r.total > T.partaiMaks)       langgar.push('volume ' + r.total);
  if (r.kalahBeruntun >= 3)         langgar.push('tilt ' + r.kalahBeruntun + ' kalah beruntun');
  if (rushPercobaan === 0)          langgar.push('Rush tidak dikerjakan');
  if (rushPercobaan > 0 && rushSkor < T.rushSkor) langgar.push('skor Rush ' + rushSkor + ' < ' + T.rushSkor);

  const row = [
    tanggal, status, rushPercobaan, rushSkor, rBest ? rBest.score : '',
    taktikTinggi, rekorBaru ? 'YA' : '',
    stats && stats.chess_rapid && stats.chess_rapid.last ? stats.chess_rapid.last.rating : '',
    stats && stats.chess_blitz && stats.chess_blitz.last ? stats.chess_blitz.last.rating : '',
    r.total, r.rapid, r.blitz, r.bullet, r.menang, r.kalah, r.seri,
    r.kalahBeruntun, r.akurasi || '', kepatuhan, langgar.join('; '),
  ];

  // Satu baris per tanggal: kalau hari ini sudah ada, timpa.
  const data = sh.getDataRange().getValues();
  let idx = -1;
  for (let i = 1; i < data.length; i++) if (data[i][0] === tanggal) { idx = i + 1; break; }
  if (idx > 0) sh.getRange(idx, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);

  return { tanggal, rushPercobaan, rushSkor, rekorBaru, kepatuhan, langgar, r, status };
}

function ringkasPartai_(partai) {
  const me = KONFIG.chessUsername.toLowerCase();
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
  const me = KONFIG.chessUsername.toLowerCase();
  const adaUrl = {};
  sh.getDataRange().getValues().slice(1).forEach(r => adaUrl[r[8]] = true);

  const seri = ['agreed', 'repetition', 'stalemate', 'insufficient', '50move', 'timevsinsufficient'];
  const baru = partai.filter(g => !adaUrl[g.url]).sort((a, b) => a.end_time - b.end_time).map(g => {
    const putih = (g.white.username || '').toLowerCase() === me;
    const sisi = putih ? g.white : g.black;
    const lawan = putih ? g.black : g.white;
    const hasil = sisi.result === 'win' ? 'M' : (seri.indexOf(sisi.result) >= 0 ? 'R' : 'K');
    const acc = g.accuracies ? (putih ? g.accuracies.white : g.accuracies.black) : '';
    return [
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
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
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
  log_('OK', 'Backfill selesai: ' + n + ' partai dari ' + arsip.length + ' bulan.');
}

// ═══════════════════ LAPORAN ═══════════════════
function kirimLaporan_(tanggal, b) {
  const T = KONFIG.target;
  const lulus = x => x ? '\u2713' : '\u2717';

  const badan = [
    'LAPORAN HARIAN — ' + KONFIG.namaAnak + ' — ' + tanggal,
    '',
    'Kepatuhan terverifikasi: ' + b.kepatuhan + '%',
    '',
    lulus(b.rushPercobaan >= T.rushPercobaan) + ' Puzzle Rush: ' + b.rushPercobaan + '/' + T.rushPercobaan +
      ' percobaan (skor ' + b.rushSkor + ', target ' + T.rushSkor + ')',
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
    'Angka di atas HANYA mencakup yang bisa diverifikasi API Chess.com.',
    'Puzzle di menu Puzzles biasa, pelajaran, dan latihan papan fisik TIDAK terlihat di sini.',
    'Untuk itu, tanya anaknya. Dasbor tidak menggantikan percakapan.',
  ].filter(x => x !== '').join('\n');

  kirimEmail_('[RTGM] ' + KONFIG.namaAnak + ' — ' + tanggal + ' — kepatuhan ' + b.kepatuhan + '%', badan);
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
function doGet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_HARIAN);
  const data = sh ? sh.getDataRange().getValues() : [];
  const snaps = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    snaps[r[0]] = {
      rushAttempts: r[2] || 0, rushScore: r[3] || 0,
      tacticsHigh: r[5] || null, rapid: r[7] || null, blitz: r[8] || null,
      status: r[1], kepatuhan: r[18], langgar: r[19],
    };
  }
  let stats = null, games = [];
  try { stats = ambilStats_(); } catch (e) { /* pakai catatan saja */ }
  try {
    const arsip = ambilArsip_();
    if (arsip.length) games = ambilBulan_(arsip[arsip.length - 1]);
  } catch (e) { /* abaikan */ }

  return ContentService
    .createTextOutput(JSON.stringify({
      user: KONFIG.chessUsername.toLowerCase(),
      nama: KONFIG.namaAnak,
      stats: stats, games: games, snaps: snaps,
      diambil: new Date().toISOString(),
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════ UJI MANUAL ═══════════════════
/** Jalankan ini dulu sebelum pasang(), untuk memastikan API bisa diakses. */
function ujiKoneksi() {
  try {
    const s = ambilStats_();
    const rush = s.puzzle_rush && s.puzzle_rush.daily ? s.puzzle_rush.daily : null;
    const msg = 'BERHASIL.\n' +
      'Rapid: ' + (s.chess_rapid && s.chess_rapid.last ? s.chess_rapid.last.rating : '-') + '\n' +
      'Taktik tertinggi: ' + (s.tactics && s.tactics.highest ? s.tactics.highest.rating : '-') + '\n' +
      'Puzzle Rush hari ini: ' + (rush ? rush.total_attempts + ' percobaan, skor ' + rush.score : 'belum ada') + '\n' +
      'Puzzle Rush terbaik: ' + (s.puzzle_rush && s.puzzle_rush.best ? s.puzzle_rush.best.score : '-');
    console.log(msg);
    return msg;
  } catch (e) {
    console.log('GAGAL: ' + e.message);
    return 'GAGAL: ' + e.message;
  }
}
