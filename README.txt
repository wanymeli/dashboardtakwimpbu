======================================================================
  DASHBOARD LAPORAN PELAKSANAAN AKTIVITI
  POLITEKNIK BALIK PULAU (PBU) — TAHUN 2026
  README.txt — Arahan Penggunaan & Maklumat Teknikal
======================================================================

VERSI    : 2.0
TARIKH   : 2026
PLATFORM : Pelayar Web Moden (tanpa pelayan)
TEMPOH DATA: Q1 2026 (Januari – Mac)

----------------------------------------------------------------------
A. KANDUNGAN PAKEJ
----------------------------------------------------------------------

  index.html   — Halaman utama dashboard (struktur HTML)
  style.css    — Reka bentuk dan susun atur (CSS)
  script.js    — Logik pengiraan, carta dan interaktiviti (JavaScript)
  README.txt   — Fail arahan ini

  FAIL DATA UNTUK DIIMPORT:
    1. AKTIVITI JANUARI 2026.xlsx
    2. AKTIVITI FEBRUARI 2026.xlsx
    3. AKTIVITI MAC 2026.xlsx

  NOTA: Data sebenar MESTI diimport daripada fail Excel asal.
  Tiada data simulasi disediakan dalam versi ini.

----------------------------------------------------------------------
B. KEPERLUAN SISTEM
----------------------------------------------------------------------

  • Pelayar web moden yang disokong:
      - Google Chrome 100+
      - Microsoft Edge 100+
      - Mozilla Firefox 100+
      - Safari 15+

  • SAMBUNGAN INTERNET DIPERLUKAN pada kali pertama kerana dashboard
    menggunakan pustaka CDN berikut:
      - Chart.js       : https://cdn.jsdelivr.net/npm/chart.js@4.4.0/
      - SheetJS (xlsx) : https://cdn.jsdelivr.net/npm/xlsx@0.18.5/
      - jsPDF          : https://cdnjs.cloudflare.com/ajax/libs/jspdf/
      - jsPDF-AutoTable: https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/

  • Tiada pelayan web, pangkalan data atau pemasangan perisian
    tambahan diperlukan.

----------------------------------------------------------------------
C. CARA MENJALANKAN DASHBOARD
----------------------------------------------------------------------

  LANGKAH 1: Pastikan ketiga-tiga fail (index.html, style.css,
             script.js) berada dalam FOLDER YANG SAMA.

  LANGKAH 2: Buka fail "index.html" menggunakan pelayar web.
             Cara: Klik dua kali pada fail index.html
             ATAU  Seret fail index.html ke tetingkap pelayar.

  LANGKAH 3: Dashboard akan dimuatkan dalam keadaan kosong (TIADA DATA).
             Notis import akan dipaparkan pada bahagian atas.

  LANGKAH 4: Import data Excel untuk mengaktifkan semua paparan.
             Ikut arahan di bahagian D di bawah.

----------------------------------------------------------------------
D. CARA MENGIMPORT DATA EXCEL
----------------------------------------------------------------------

  FORMAT FAIL YANG DISOKONG:
  Dashboard ini direka untuk membaca fail Excel bulanan PBU dengan
  format lajur asal:

    | Hari | Tarikh | Program / Aktiviti | Peserta |
    | Jabatan / Unit | Pegawai Urusetia (PIC) |
    | Catatan / Justifikasi | Status Pelaksanaan |

  Variasi nama lajur juga disokong secara automatik.

  KAEDAH IMPORT:

  LANGKAH 1: Dari menu kiri, klik "Pengurusan Data".

  LANGKAH 2: Dalam bahagian "Import Fail Excel Aktiviti", anda boleh:
     (a) Klik butang "Pilih Fail Excel" dan pilih satu atau lebih fail
     (b) Seret terus fail Excel ke dalam kotak import (drag & drop)

  LANGKAH 3: Import BERBILANG FAIL sekaligus:
     - Pilih ketiga-tiga fail (Jan, Feb, Mac) serentak
     - ATAU import satu persatu (data akan digabungkan)

  LANGKAH 4: Log import akan dipaparkan secara automatik:
     ✓ Bilangan rekod dibaca dari setiap fail
     ✓ Amaran pendua (jika ada)
     ✓ Amaran percanggahan status-catatan
     ✓ Rekod dengan medan kosong

  LANGKAH 5: Setelah import berjaya, semua KPI, graf dan jadual
             akan dikemas kini secara automatik.

  NOTA PENTING:
  • Sistem membaca SEMUA helaian dalam setiap fail Excel.
  • Baris kosong dan baris tajuk berulang akan ditolak secara automatik.
  • Format tarikh dd/mm/yyyy, dd-mm-yyyy dan nombor serial Excel
    semua disokong.
  • Status seperti "SELESAI", "Selesai", "selesai" semua diterima.

----------------------------------------------------------------------
E. HALAMAN-HALAMAN DASHBOARD
----------------------------------------------------------------------

  1. DASHBOARD UTAMA
     ---------------
     • 7 KPI utama berdasarkan data sebenar.
     • Graf perbandingan aktiviti Jan–Mac 2026.
     • Pecahan status (Selesai / Ditunda / Lain).
     • Trend bilangan aktiviti bulanan.
     • Prestasi mengikut Jabatan/Unit.
     • 10 rekod terkini.

  2. LAPORAN BULANAN
     ----------------
     • Pilih Jan, Feb atau Mac untuk paparan terperinci.
     • April–Disember dipaparkan sebagai "Data Belum Dibekalkan"
       (tidak dianggap sifar aktiviti).
     • Senarai aktiviti ditunda dengan catatan penundaan.
     • Eksport PDF laporan bulanan.

  3. LAPORAN SUKU TAHUNAN
     ---------------------
     • Fokus pada Q1 2026 (Januari – Mac).
     • Perbandingan 3 bulan dalam Q1.
     • Jadual pencapaian bulanan.
     • Q2, Q3, Q4 dipaparkan dengan status "Data Belum Dibekalkan".

  4. ANALISIS JABATAN
     -----------------
     • Perbandingan Selesai vs Ditunda mengikut Jabatan/Unit.
     • Nota bahawa bilangan aktiviti lebih tinggi tidak semestinya
       menunjukkan prestasi lebih baik.

  5. AKTIVITI DITUNDA
     -----------------
     • Senarai lengkap semua aktiviti berstatus Ditunda.
     • Catatan/justifikasi penundaan daripada rekod asal.
     • Cadangan tindakan susulan berdasarkan catatan.
     • Eksport PDF laporan ditunda.

  6. LOG KUALITI DATA
     -----------------
     • Rekod asal vs rekod selepas gabungan.
     • Rekod dengan percanggahan status vs catatan.
     • Rekod dengan medan kosong.
     • Rekod dengan tarikh meragukan.
     • Senarai pembetulan yang dibuat.
     PENTING: Rekod isu TIDAK dipadam — dipaparkan untuk pengesahan.

  7. SENARAI AKTIVITI
     -----------------
     • Semua rekod dengan penapis: Jabatan, Status, Suku, Kata Kunci.
     • Paparan butiran penuh setiap rekod.
     • Edit dan padam rekod.
     • Eksport Excel dan PDF.

  8. PENGURUSAN DATA
     ----------------
     • Import fail Excel (multi-fail, drag & drop).
     • Eksport pelbagai format laporan.
     • Tambah/edit rekod secara manual.
     • Padam semua data.

----------------------------------------------------------------------
F. CARA MENGEKSPORT LAPORAN
----------------------------------------------------------------------

  EKSPORT PDF — LAPORAN PENGURUSAN (LENGKAP):
  1. Pergi ke "Pengurusan Data".
  2. Klik "Laporan Pengurusan Q1 2026 (PDF)".
  3. PDF merangkumi: KPI, analisis bulanan, analisis jabatan,
     senarai ditunda, isu kualiti data.

  EKSPORT PDF — LAPORAN BULANAN:
  1. Pergi ke "Laporan Bulanan".
  2. Pilih bulan (atau "Semua Bulan").
  3. Klik "PDF Laporan".

  EKSPORT PDF — LAPORAN DITUNDA:
  1. Pergi ke "Aktiviti Ditunda".
  2. Tapis jika perlu.
  3. Klik "PDF Laporan Ditunda".

  EKSPORT EXCEL:
  Dalam "Pengurusan Data":
  • "Semua Aktiviti"     — dataset gabungan dengan medan analisis
  • "Laporan Bulanan"    — ringkasan Jan, Feb, Mac
  • "Laporan Suku Q1"    — ringkasan Q1 dengan status Q2-Q4
  • "Analisis Jabatan"   — prestasi mengikut jabatan
  • "Senarai Ditunda"    — rekod ditunda sahaja

  Dalam "Senarai Aktiviti":
  • "Eksport Excel"      — berdasarkan penapis semasa

----------------------------------------------------------------------
G. PENYIMPANAN DATA (LOCALSTORAGE)
----------------------------------------------------------------------

  ⚠ PENTING: Data disimpan dalam LocalStorage pelayar sahaja.

  • Data TIDAK dihantar ke mana-mana pelayan.
  • Data HANYA tersedia pada pelayar dan peranti yang sama.
  • Jika pelayar dikosongkan (clear cache), data AKAN hilang.

  AMALAN BAIK:
  ✓ Eksport data sebagai fail Excel selepas setiap import.
  ✓ Simpan fail Excel di folder selamat sebagai sandaran.
  ✓ Gunakan komputer yang sama untuk memantau dashboard.

----------------------------------------------------------------------
H. FORMULA PENGIRAAN
----------------------------------------------------------------------

  PERATUS SELESAI:
    = (Bilangan Selesai ÷ Jumlah Semua Rekod) × 100%
    = "N/A" jika tiada rekod

  PERATUS DITUNDA:
    = (Bilangan Ditunda ÷ Jumlah Semua Rekod) × 100%

  NOTA KAWALAN KETEPATAN:
  • Semua peratus dikira daripada rekod yang ada dalam fail Excel.
  • Rekod yang tiada status tidak dikira sebagai Selesai atau Ditunda.
  • Status berdasarkan rekod sumber — tertakluk pengesahan pegawai.
  • Bulan yang tiada data (Apr–Dis) TIDAK dianggap sifar aktiviti.
  • Tidak dibuat kesimpulan prestasi tahunan daripada data Q1 sahaja.

----------------------------------------------------------------------
I. PEMETAAN LAJUR EXCEL (RUJUKAN TEKNIKAL)
----------------------------------------------------------------------

  Sistem mengenal pasti lajur secara automatik berdasarkan nama:

  Medan Hari:    "Hari", "Day", "Hari Pelaksanaan"
  Medan Tarikh:  "Tarikh", "Date", "Tarikh Pelaksanaan", "Tarikh Program"
  Medan Nama:    "Program / Aktiviti", "Program", "Aktiviti",
                 "Nama Aktiviti", "Tajuk Program"
  Medan Peserta: "Peserta", "Bilangan Peserta", "Penerima"
  Medan Jabatan: "Jabatan / Unit", "Jabatan", "Unit", "Bahagian"
  Medan PIC:     "Pegawai Urusetia (PIC)", "Pegawai Urusetia",
                 "PIC", "Pegawai Bertanggungjawab"
  Medan Catatan: "Catatan / Justifikasi", "Catatan", "Justifikasi"
  Medan Status:  "Status Pelaksanaan", "Status", "Status Aktiviti"

  Variasi huruf besar/kecil dan spasi tambahan diterima.

----------------------------------------------------------------------
J. PENYELESAIAN MASALAH
----------------------------------------------------------------------

  MASALAH: Import Excel tidak menunjukkan data
  PENYELESAIAN:
  • Pastikan fail tidak dilindungi kata laluan
  • Semak bahawa lajur mengandungi sekurang-kurangnya nama aktiviti
    atau tarikh
  • Cuba import semula

  MASALAH: Tarikh tidak dikenali
  PENYELESAIAN:
  • Format tarikh yang disokong: dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd
  • Tarikh Excel serial number turut disokong
  • Rekod akan dikekalkan walaupun tarikh tidak dikenali

  MASALAH: Status tidak dipaparkan dengan betul
  PENYELESAIAN:
  • Nilai diterima: SELESAI, Selesai, selesai, DONE, COMPLETED
  • Nilai diterima: DITUNDA, Ditunda, ditunda, TANGGUH, POSTPONED
  • Status lain dipaparkan sebagaimana dalam fail

  MASALAH: Graf tidak dipaparkan
  PENYELESAIAN: Semak sambungan internet (CDN Chart.js perlu diakses)

  MASALAH: Data hilang selepas tutup pelayar
  PENYELESAIAN: Import semula dari fail Excel asal
                ATAU eksport dahulu sebelum tutup pelayar

----------------------------------------------------------------------
K. NOTA TAMBAHAN
----------------------------------------------------------------------

  • Logo rasmi PBU boleh ditambah dengan mengedit index.html,
    tukar <div class="logo-placeholder">PBU</div> kepada
    <img src="logo-pbu.png" alt="PBU" style="height:42px;border-radius:8px">

  • Untuk pertanyaan teknikal, hubungi Unit ICT Politeknik Balik Pulau.

  • Laporan ini HANYA meliputi Q1 2026 (Januari – Mac) dan tidak
    mewakili keseluruhan tahun 2026.

  • Data Q2–Q4 2026 akan dipaparkan apabila fail Excel berkenaan
    diimport ke dalam sistem.

----------------------------------------------------------------------
  Politeknik Balik Pulau | Laporan Pelaksanaan Aktiviti 2026
  Dashboard HTML/CSS/JavaScript v2.0
  Tempoh Data: Januari – Mac 2026 (Q1)
======================================================================
