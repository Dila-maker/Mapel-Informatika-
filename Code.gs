/**
 * GENERATOR PROMPT PRESENTASI PEMBELAJARAN INTERAKTIF
 * Google Apps Script (V8 / ES6+) - tanpa library eksternal.
 *
 * Sheet "Input"  : kolom A = label, kolom B = isian guru (baris 1-20)
 * Sheet "Prompt" : sel A1 = prompt raksasa hasil rangkaian script
 */

// ---------- KONSTANTA ----------
const SHEET_INPUT = 'Input';
const SHEET_PROMPT = 'Prompt';
const JUMLAH_BARIS_INPUT = 20;

const LABEL_INPUT = [
  'MAPEL', 'NAMA MATERI',
  'TUJUAN 1', 'TUJUAN 2', 'TUJUAN 3', 'TUJUAN 4', 'TUJUAN 5',
  'SUBTOPIK 1', 'SUBTOPIK 2', 'SUBTOPIK 3', 'SUBTOPIK 4', 'SUBTOPIK 5',
  'WARNA 1', 'WARNA 2', 'WARNA 3',
  'IKON 1', 'IKON 2', 'IKON 3', 'IKON 4', 'IKON 5'
];

// ---------- MENU ----------
/** Dijalankan otomatis saat spreadsheet dibuka: membuat menu kustom. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Menu Prompt')
    .addItem('🚀 Generate Prompt', 'generatePrompt')
    .addItem('📋 Copy Prompt ke Clipboard', 'copyPrompt')
    .addToUi();
}

// ---------- SETUP (opsional, jalankan sekali dari editor) ----------
/** Membuat sheet "Input" & "Prompt" beserta label kolom A secara otomatis. */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet Input
  let input = ss.getSheetByName(SHEET_INPUT);
  if (!input) input = ss.insertSheet(SHEET_INPUT);
  input.getRange(1, 1, JUMLAH_BARIS_INPUT, 1)
    .setValues(LABEL_INPUT.map(label => [label]))
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  input.setColumnWidth(1, 160);
  input.setColumnWidth(2, 520);
  input.getRange(1, 2, JUMLAH_BARIS_INPUT, 1).setWrap(true);

  // Sheet Prompt
  let prompt = ss.getSheetByName(SHEET_PROMPT);
  if (!prompt) prompt = ss.insertSheet(SHEET_PROMPT);
  prompt.setColumnWidth(1, 800);
  prompt.getRange('A1').setWrap(true).setVerticalAlignment('top');

  SpreadsheetApp.getUi().alert('✅ Sheet "Input" dan "Prompt" siap dipakai.');
}

// ---------- GENERATE PROMPT ----------
/** Mengambil input guru, merangkai prompt, lalu menyimpannya ke Prompt!A1. */
function generatePrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Cek keberadaan sheet
  const input = ss.getSheetByName(SHEET_INPUT);
  if (!input) {
    ui.alert('❌ Sheet "' + SHEET_INPUT + '" tidak ditemukan. Jalankan setupSheets() atau buat sheet secara manual.');
    return;
  }
  let promptSheet = ss.getSheetByName(SHEET_PROMPT);
  if (!promptSheet) promptSheet = ss.insertSheet(SHEET_PROMPT);

  // Ambil data kolom B baris 1-20
  const nilai = input.getRange(1, 2, JUMLAH_BARIS_INPUT, 1).getValues()
    .map(baris => String(baris[0] === null || baris[0] === undefined ? '' : baris[0]).trim());

  // Validasi minimal: mapel & nama materi wajib diisi
  if (!nilai[0] || !nilai[1]) {
    ui.alert('⚠️ MAPEL (B1) dan NAMA MATERI (B2) wajib diisi dulu.');
    return;
  }

  // Isian kosong diganti tanda strip agar prompt tetap rapi
  const v = nilai.map(x => x || '-');

  const hasil = buatPrompt({
    mapel: v[0], materi: v[1],
    tujuan1: v[2], tujuan2: v[3], tujuan3: v[4], tujuan4: v[5], tujuan5: v[6],
    sub1: v[7], sub2: v[8], sub3: v[9], sub4: v[10], sub5: v[11],
    warna1: v[12], warna2: v[13], warna3: v[14],
    ikon1: v[15], ikon2: v[16], ikon3: v[17], ikon4: v[18], ikon5: v[19]
  });

  // Simpan ke Prompt!A1 + format tampilan
  promptSheet.setColumnWidth(1, 800);
  const sel = promptSheet.getRange('A1');
  sel.setValue(hasil).setWrap(true).setVerticalAlignment('top');
  try {
    promptSheet.autoResizeRows(1, 1);
  } catch (e) {
    // Abaikan jika auto-resize gagal (teks sangat panjang)
  }

  ss.setActiveSheet(promptSheet);
  ui.alert('✅ Prompt berhasil dibuat (' + hasil.length + ' karakter).\n\nLanjutkan dengan menu: ⚡ Menu Prompt > 📋 Copy Prompt ke Clipboard.');
}

// ---------- COPY KE CLIPBOARD ----------
/** Membuka dialog kecil yang menyalin isi Prompt!A1 ke clipboard. */
function copyPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const promptSheet = ss.getSheetByName(SHEET_PROMPT);
  if (!promptSheet) {
    ui.alert('❌ Sheet "' + SHEET_PROMPT + '" tidak ditemukan.');
    return;
  }
  const teks = String(promptSheet.getRange('A1').getValue() || '');
  if (!teks) {
    ui.alert('⚠️ Sel A1 pada sheet "Prompt" masih kosong. Klik "🚀 Generate Prompt" dulu.');
    return;
  }

  // Aman disisipkan ke dalam <script>: escape karakter "<" dan pemisah baris
  const teksAman = JSON.stringify(teks)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');

  const html =
    '<div style="font-family:Arial,sans-serif;padding:8px;">' +
    '<p id="status" style="font-size:14px;">⏳ Menyalin prompt...</p>' +
    '<textarea id="kotak" style="width:100%;height:120px;font-size:12px;" readonly></textarea>' +
    '<p><button id="tombol" style="padding:8px 16px;font-size:14px;cursor:pointer;">📋 Salin Lagi</button> ' +
    '<button id="tutup" style="padding:8px 16px;font-size:14px;cursor:pointer;">Tutup</button></p>' +
    '</div>' +
    '<script>' +
    'const teks = ' + teksAman + ';' +
    'const status = document.getElementById("status");' +
    'const kotak = document.getElementById("kotak");' +
    'kotak.value = teks;' +
    // Cadangan jika navigator.clipboard diblokir browser
    'function salinCadangan() {' +
    '  kotak.focus(); kotak.select();' +
    '  try { return document.execCommand("copy"); } catch (e) { return false; }' +
    '}' +
    'async function salin() {' +
    '  try {' +
    '    await navigator.clipboard.writeText(teks);' +
    '    status.textContent = "✅ Prompt berhasil disalin! Silakan paste (Ctrl+V) ke ChatGPT/Claude/Gemini.";' +
    '  } catch (err) {' +
    '    if (salinCadangan()) {' +
    '      status.textContent = "✅ Prompt berhasil disalin! Silakan paste (Ctrl+V) ke ChatGPT/Claude/Gemini.";' +
    '    } else {' +
    '      status.textContent = "⚠️ Gagal otomatis. Klik kotak teks, tekan Ctrl+A lalu Ctrl+C.";' +
    '    }' +
    '  }' +
    '}' +
    'document.getElementById("tombol").addEventListener("click", salin);' +
    'document.getElementById("tutup").addEventListener("click", function () { google.script.host.close(); });' +
    'salin();' +
    '</script>';

  const dialog = HtmlService.createHtmlOutput(html).setWidth(520).setHeight(300);
  ui.showModalDialog(dialog, '📋 Copy Prompt');
}

// ---------- TEMPLATE PROMPT RAKSASA ----------
/** Merangkai template prompt dengan template literal (backtick) + variabel. */
function buatPrompt(d) {
  const { mapel, materi, tujuan1, tujuan2, tujuan3, tujuan4, tujuan5,
          sub1, sub2, sub3, sub4, sub5, warna1, warna2, warna3,
          ikon1, ikon2, ikon3, ikon4, ikon5 } = d;

  return `PROMPT PEMBUATAN WEBSITE PRESENTASI PEMBELAJARAN INTERAKTIF ${mapel} KELAS VII
(Materi: ${materi})

---

IDENTITAS
· Sekolah: SMPN 1 Tinggimoncong
· Mapel: ${mapel}
· Kelas: VII
· Guru: Syarifuddin, S. Pd., M. Si.
· Materi: ${materi}
· Alokasi: ±80 menit

TUJUAN PEMBELAJARAN
1. ${tujuan1}
2. ${tujuan2}
3. ${tujuan3}
4. ${tujuan4}
5. ${tujuan5}

---

KONSEP UTAMA
Website presentasi pembelajaran interaktif untuk guru mengajar
di depan kelas. BUKAN LKPD, BUKAN ujian.

Prinsip:
· Tanpa input nama siswa, tanpa penilaian, tanpa sertifikat.
· 20% penjelasan + 80% aktivitas interaktif.
· 1 slide = 1 konsep utama (materi singkat).
· Setiap topik diselingi game edukasi.
· Setiap 2–3 slide materi disisipkan contoh soal interaktif.
· Di akhir pembelajaran: latihan gamifikasi.
· Fokus pengalaman belajar menyenangkan & bermakna.

Mode Presentasi:
· Fullscreen, ditampilkan guru di depan kelas.
· Navigasi slide besar: ⬅️ Sebelumnya | ➡️ Berikutnya.
· Indikator slide (● ● ○ ○ ○).
· Animasi transisi halus.
· Dukungan keyboard (← →).

---

PRINSIP STRUKTUR BERAGAM (WAJIB — ANTI-BOSAN)
Struktur alur presentasi TIDAK HARUS memakai pola
"Pos 1, Pos 2, Pos 3..." secara seragam.
Gunakan STRUKTUR BERCAMPUR (mixed structure) agar
setiap bagian terasa segar dan tidak monoton.

Contoh kerangka alur yang bisa dipilih/dikombinasikan:
· 🚪 Gerbang Pembuka (Welcome)
· 🗺️ Peta Petualangan (opsional)
· 📖 Bab Cerita / Misi
· 🎯 Tantangan Kilat (mini quiz)
· 🔬 Laboratorium / Simulasi
· 🧩 Puzzle / Misteri
· 🎬 Studio Kreatif (buat karya)
· ⚔️ Arena Battle (2 tim)
· 🏆 Boss Battle / Ujian Akhir
· 🎉 Panggung Perayaan (rangkuman & refleksi)

Variasikan elemen berikut antar-bagian:
· Layout slide (kiri-kanan, atas-bawah, center, grid)
· Cara interaksi (klik, drag, tebak, gambar, suara, gerak)
· Media (emoji, ilustrasi CSS, animasi, teks besar, kartu)
· Nada bahasa (petualangan, misteri, tantangan, kolaborasi)

---

TEMA KHUSUS MATERI (WAJIB)
Website punya tema visual khusus ${materi}, bukan generik.
· Palet warna utama: ${warna1}, ${warna2}, ${warna3}
· Ikon khas: ${ikon1}, ${ikon2}, ${ikon3}, ${ikon4}, ${ikon5}
· Ornamen latar bertema materi + animasi ringan
· Setiap bagian punya warna & ikon berbeda

---

STRUKTUR HALAMAN (CONTOH BERCAMPUR)

1. PEMBUKA — "Gerbang Petualangan"
   - Judul: 🌟 PETUALANGAN ${materi} 🌟
   - Subjudul: Presentasi Interaktif ${mapel} Kelas VII
   - Info: SMPN 1 Tinggimoncong, Guru
   - Ilustrasi animasi ikon bertema materi
   - Tombol: 🚀 MULAI PETUALANGAN

2. PETA / MENU UTAMA
   - Peta visual berisi 5–7 "wilayah"
   - Contoh penamaan:
     🏠 Bab 1 — ${sub1}
     🔬 Lab — ${sub2}
     🧩 Misteri — ${sub3}
     ⚔️ Arena — ${sub4}
     🎬 Studio — ${sub5}
     🏆 Boss — Tantangan Akhir
   - Progress bar 0% ━━━━━ 100%

3. BAB 1 — ${sub1} (±10–12 menit)
   Struktur boleh: Materi → Game → Soal → Game

4. BAB 2 — ${sub2} (±10–12 menit)
   Struktur boleh: Game → Materi → Soal

5. BAB 3 — ${sub3} (±10–12 menit)
   Struktur boleh: Materi → Simulasi → Soal → Puzzle

6. BAB 4 — ${sub4} (±10–12 menit)
   Struktur boleh: Cerita → Materi → Battle → Soal

7. BAB 5 — ${sub5} (±10–12 menit)
   Struktur boleh: Materi → Kreatif → Soal

8. BOSS BATTLE / TANTANGAN AKHIR (±15 menit)
   Kompilasi 6–8 game gamifikasi campuran.

9. PENUTUP — "Panggung Perayaan"
   - 🎉 PETUALANGAN SELESAI!
   - Rangkuman 5–6 poin utama materi
   - Refleksi bersama
   - Tombol: 🔁 ULANGI | 🏁 SELESAI

---

DAFTAR 60 GAME PEMBELAJARAN (PILIHAN REFERENSI)
A. GAME KUIS & TANYA-JAWAB: Kuis rebutan, Kahoot, Jeopardy, Millionaire,
   Family Feud, Spin wheel, Truth or dare, Hot seat.
B. GAME PAPAN & KARTU: Ular tangga, Monopoli, Bingo, Domino, Flash card,
   Memory card, UNO, Scrabble.
C. GAME TEBAK & MISTERI: Tebak gambar, Tebak tokoh, Tebak suara, Charades,
   Pictionary, Riddle, Mystery box.
D. GAME PETUALANGAN: Escape room, Scavenger hunt, Treasure hunt, Quest RPG,
   Boss battle, Level-based.
E. GAME SIMULASI: Role-playing, Debat, Sidang, Pasar, PBB, Model UN.
F. GAME FISIK: Relay soal, Four corners, Simon says, Human knot,
   Treasure race, Ball toss.
G. GAME KOLABORATIF: Jigsaw, Think-pair-share, Team building,
   Escape classroom, Debat kelompok, Gallery walk.
H. GAME DIGITAL: Minecraft Edu, Roblox Edu, Scratch, Prodigy, Duolingo,
   PhET, Wordwall, Blooket.
I. GAME KREATIF: Puzzle kolaboratif, Komik, Lagu/Rap, Drama mini, Stop motion.

PENYESUAIAN GAME DENGAN MATERI ${materi}
· Bab 1 — ${sub1}: Game: [No & Nama Game], Alasan: [...]
· Bab 2 — ${sub2}: Game: [No & Nama Game], Alasan: [...]
· Bab 3 — ${sub3}: Game: [No & Nama Game], Alasan: [...]
· Bab 4 — ${sub4}: Game: [No & Nama Game], Alasan: [...]
· Bab 5 — ${sub5}: Game: [No & Nama Game], Alasan: [...]
· Boss Battle — 6–8 game campuran: [No & Nama Game]

---

MODE PRESENTASI INTERAKTIF
· Setiap bab/slide = satu layar besar fullscreen
· Tombol navigasi besar: ⬅️ Sebelumnya | ➡️ Berikutnya
· Indikator slide di bawah
· Animasi transisi halus (boleh berbeda per bab)
· Tombol 🔍 zoom visual
· Keyboard ← → untuk navigasi
· Tombol 🎮 untuk masuk mode game
· Tanpa input nama, tanpa nilai, tanpa sertifikat

---

ALOKASI WAKTU 80 MENIT
Pembukaan & peta | 3 menit
Bab 1 (materi + game + soal) | 10 menit
Bab 2 (materi + game + soal) | 10 menit
Bab 3 (materi + game + soal) | 12 menit
Bab 4 (materi + game + soal) | 12 menit
Bab 5 (materi + game + soal) | 12 menit
Boss Battle (6–8 game gamifikasi) | 15 menit
Rangkuman & Refleksi | 6 menit
Total | 80 menit

---

DESAIN VISUAL
· Modern, ceria, penuh warna
· Ramah siswa SMP
· Responsif (HP/tablet/PC)
· Banyak ikon & ilustrasi bertema materi
· Tombol besar & menarik
· Animasi ringan
· Efek suara ringan (opsional)
· Progress bar di setiap bab
· Warna berbeda per bab
· Layout bervariasi antar-bab
· Feedback jawaban berwarna

---

FITUR WAJIB
✅ HTML+CSS+JS satu file, langsung jalan, tanpa server
✅ Tanpa login, tanpa database
✅ Tanpa input nama, tanpa nilai, tanpa sertifikat
✅ Responsif (HP/tablet/PC)
✅ Mode presentasi fullscreen untuk guru
✅ Navigasi slide jelas + keyboard (← →)
✅ Materi singkat, padat, jelas
✅ Struktur bercampur (bukan hanya "Pos 1, 2, 3")
✅ Nama bab bervariasi (bab/lab/misteri/arena/studio)
✅ Setiap bab: minimal 1 materi + 1 game + 1 contoh soal
✅ Boss Battle: 6–8 game gamifikasi bervariasi
✅ Game dipilih dari Daftar 60 Game Pembelajaran
✅ Setiap game disesuaikan dengan materi & tujuan
✅ Feedback langsung setiap aksi
✅ Animasi ringan & menarik
✅ Tema visual khusus materi (bukan generik)

---

PERINTAH ANTI-ERROR (WAJIB DIPATUHI)
1. Bebas error di browser (Chrome/Edge/Firefox/Safari), tanpa console error.
2. Tidak pakai sintaks deprecated.
3. Fungsi didefinisikan sebelum dipanggil / hoisting benar.
4. Tidak ada variabel undefined/null tanpa pengecekan.
5. Event listener pakai addEventListener.
6. Elemen DOM dicek if (element) sebelum diakses.
7. Semua tag HTML ditutup benar.
8. CSS valid, tidak ada typo properti.
9. Escape karakter khusus dengan benar.
10. Beri komentar singkat per bagian kode.
11. Semua tombol navigasi berfungsi, tidak ada yang mati.
12. Tidak ada alert() bawaan; gunakan modal kustom.
13. Animasi tidak mengganggu performa.
14. Emoji & ikon dirender konsisten di semua browser.
15. Tidak pakai library eksternal berlebihan; utamakan HTML/CSS/JS murni.

Buat seluruh kode HTML + CSS + JavaScript lengkap dalam
SATU FILE (index.html) yang langsung dapat dijalankan,
tanpa error, tampilan sangat menarik, berwarna-warni,
bertema ${materi}, dengan banyak gambar dan ikon.

---

CATATAN MATERI
· Setiap konsep disajikan singkat (1 slide = 1 konsep).
· Gunakan analogi sederhana & contoh nyata.
· Sisipkan game agar siswa tidak bosan.
· Sertakan contoh soal interaktif untuk dibahas bersama guru.
· Boss Battle menguji seluruh materi secara menyenangkan.
· Sertakan refleksi bersama di akhir.
· Pilih game dari Daftar 60 Game Pembelajaran yang paling sesuai.
· Variasikan struktur antar-bab agar tidak monoton.`;
}
