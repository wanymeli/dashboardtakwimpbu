/* =====================================================
   DASHBOARD LAPORAN PELAKSANAAN AKTIVITI PBU 2026
   script.js  — Logik Utama v2.0
   =====================================================
   Sumber data: Fail Excel bulanan yang diimport melalui
   antara muka pengguna. Tiada data simulasi digunakan.
   Data sebenar diimport dari fail:
     - 1. AKTIVITI JANUARI 2026.xlsx
     - 2. AKTIVITI FEBRUARI 2026.xlsx
     - 3. AKTIVITI MAC 2026.xlsx
   ===================================================== */

'use strict';

/* =====================================================
   1. STATE APLIKASI
   ===================================================== */
let APP = {
  data: [],
  hasData: false,
  lastUpdated: null,
  filesImported: 0,
  charts: {},
  currentPage: 1,
  pageSize: 15,
  qualityLog: {
    totalRaw: 0,
    totalAfterMerge: 0,
    emptyFields: [],
    statusIssues: [],
    duplicates: [],
    corrections: [],
    dateIssues: [],
  },
};

const MONTH_NAMES = ['Januari','Februari','Mac','April','Mei','Jun',
                     'Julai','Ogos','September','Oktober','November','Disember'];

/* =====================================================
   2. PEMETAAN LAJUR EXCEL
   Sistem menyokong pelbagai variasi nama lajur dari
   fail Excel asal PBU.
   ===================================================== */
const COLUMN_MAP = {
  hari:    ['hari','day','hari pelaksanaan'],
  tarikh:  ['tarikh','date','tarikh pelaksanaan','tarikh program','tarikh aktiviti'],
  nama:    ['program / aktiviti','program/aktiviti','program','aktiviti','nama program',
            'nama aktiviti','program aktiviti','tajuk program','tajuk aktiviti',
            'perkara','bil. / program'],
  peserta: ['peserta','bilangan peserta','bil peserta','peserta hadir','penerima',
            'bilangan penerima','bil. peserta','anggaran peserta'],
  jabatan: ['jabatan / unit','jabatan/unit','jabatan','unit','bahagian',
            'jabatan / unit bertanggungjawab','jabatan unit'],
  pic:     ['pegawai urusetia (pic)','pegawai urusetia','pic','pegawai bertanggungjawab',
            'urus setia','pegawai','pegawai pelaksana','pegawai urus setia'],
  catatan: ['catatan / justifikasi','catatan/justifikasi','catatan','justifikasi',
            'sebab','keterangan','notes','pemerhatian','ulasan'],
  status:  ['status pelaksanaan','status','status aktiviti','status program',
            'pelaksanaan','status perlaksanaan'],
};

/* =====================================================
   3. UTILITI UMUM
   ===================================================== */
function normKey(s) {
  return String(s||'').trim().toLowerCase();
}

function findColValue(row, candidates) {
  const rowKeys = Object.keys(row);
  for (const cand of candidates) {
    const found = rowKeys.find(k => normKey(k) === normKey(cand));
    if (found !== undefined) return String(row[found]||'').trim();
  }
  return '';
}

function parseDateExcel(val) {
  if (!val && val !== 0) return null;
  // Excel serial number
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m-1, d.d);
  }
  // Date object
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  // Text: dd/mm/yyyy or dd-mm-yyyy
  const s = String(val).trim();
  const ddmm = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (ddmm) {
    let [, d, m, y] = ddmm;
    if (parseInt(y) < 100) y = '20'+y;
    return new Date(parseInt(y), parseInt(m)-1, parseInt(d));
  }
  // ISO: yyyy-mm-dd
  const iso = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (iso) {
    const [, y, m, d] = iso;
    return new Date(parseInt(y), parseInt(m)-1, parseInt(d));
  }
  // General parse
  const gd = new Date(s);
  return isNaN(gd.getTime()) ? null : gd;
}

function fmtDateObj(d) {
  if (!d) return '—';
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function fmtDate(s) {
  if (!s) return '—';
  const d = new Date(s);
  return isNaN(d.getTime()) ? (s||'—') : fmtDateObj(d);
}

function getMonth(rec) {
  if (!rec.TarikhISO) return null;
  return new Date(rec.TarikhISO).getMonth() + 1;
}

function getQuarter(bulan) {
  if (bulan >= 1 && bulan <= 3)   return 'Q1';
  if (bulan >= 4 && bulan <= 6)   return 'Q2';
  if (bulan >= 7 && bulan <= 9)   return 'Q3';
  if (bulan >= 10 && bulan <= 12) return 'Q4';
  return null;
}

function badgeHtml(status) {
  const s = String(status||'').toLowerCase();
  if (s === 'selesai')     return `<span class="badge badge-selesai">Selesai</span>`;
  if (s === 'ditunda')     return `<span class="badge badge-ditunda">Ditunda</span>`;
  if (s === 'digugurkan')  return `<span class="badge badge-digugurkan">Digugurkan</span>`;
  if (s.includes('belum')) return `<span class="badge badge-belum">Belum Dibekalkan</span>`;
  return `<span class="badge badge-belum">${status||'—'}</span>`;
}

function showToast(msg, type='success') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = `toast show ${type}`;
  setTimeout(() => { t.className = 'toast'; }, 3500);
}

function calcPct(selesai, total) {
  if (!total) return 'N/A';
  return ((selesai / total) * 100).toFixed(1) + '%';
}
function calcPctNum(selesai, total) {
  if (!total) return 0;
  return parseFloat(((selesai / total) * 100).toFixed(1));
}

/* =====================================================
   4. NORMALISASI STATUS
   ===================================================== */
function normaliseStatus(raw) {
  const s = String(raw||'').trim().toUpperCase();
  if (['SELESAI','COMPLETED','DONE','TELAH DILAKSANAKAN','DILAKSANAKAN',
       'BERJAYA','OK','SIAP'].includes(s)) return 'Selesai';
  if (['DITUNDA','TANGGUH','POSTPONED','DELAYED','DEFER','TERTANGGUH',
       'BELUM SELESAI','BELUM DILAKSANAKAN','DALAM PROSES','SEDANG BERJALAN',
       'BERJALAN'].includes(s)) return 'Ditunda';
  if (['DIGUGURKAN','BATAL','DIBATALKAN','CANCELLED','CANCELED',
       'GUGUR','TIDAK DILAKSANAKAN','TIDAK JADI'].includes(s)) return 'Digugurkan';
  // Partial matches
  if (s.includes('SELESAI') || s.includes('DONE')) return 'Selesai';
  if (s.includes('TUNDA') || s.includes('TANGGUH') || s.includes('POSTPONE')) return 'Ditunda';
  if (s.includes('GUGUR') || s.includes('BATAL') || s.includes('CANCEL')) return 'Digugurkan';
  return raw ? String(raw).trim() : '';
}

/* =====================================================
   5. SEMAK KUALITI DATA
   ===================================================== */
function detectQualityIssues(rec) {
  const issues = [];
  const catatan = String(rec.Catatan||'').toLowerCase();
  const status = String(rec.StatusNormal||'').toLowerCase();

  // Status vs catatan tidak konsisten
  const ditundaKeywords = ['ditunda','tangguh','postpone','akan dilaksana','bulan hadapan',
    'bulan april','bulan mei','bulan jun','bulan julai','akan datang',
    'sedang dilaksana','dalam proses','belum','enjin','tidak dapat hadir',
    'kekangan','pertindihan','perubahan'];
  const selesaiKeywords = ['telah selesai','sudah selesai','berjaya','telah dilaksana'];

  if (status === 'selesai') {
    const hasDitundaHint = ditundaKeywords.some(k => catatan.includes(k));
    if (hasDitundaHint) {
      issues.push('Status SELESAI tetapi catatan mengandungi petanda penundaan. Perlu pengesahan.');
    }
  }
  if (status === 'ditunda') {
    const hasSelesaiHint = selesaiKeywords.some(k => catatan.includes(k));
    if (hasSelesaiHint) {
      issues.push('Status DITUNDA tetapi catatan mengandungi petanda selesai. Perlu pengesahan.');
    }
  }
  return issues;
}

function runDataQualityCheck(records) {
  const log = {
    totalRaw: 0,
    totalAfterMerge: records.length,
    emptyFields: [],
    statusIssues: [],
    duplicates: [],
    corrections: [],
    dateIssues: [],
  };

  const idSeen = {};
  records.forEach((rec, i) => {
    // Semak medan kosong utama
    const empty = [];
    if (!rec.NamaAktiviti) empty.push('Program/Aktiviti');
    if (!rec.Jabatan)      empty.push('Jabatan/Unit');
    if (!rec.TarikhISO)    empty.push('Tarikh');
    if (!rec.StatusNormal) empty.push('Status');
    if (!rec.PIC)          empty.push('PIC');
    if (empty.length > 0)  log.emptyFields.push({ rec, empty });

    // Tarikh meragukan
    if (rec.TarikhRaw && !rec.TarikhISO) {
      log.dateIssues.push({ rec, rawVal: rec.TarikhRaw });
    }
    // Tarikh bukan 2026
    if (rec.TarikhISO) {
      const yr = new Date(rec.TarikhISO).getFullYear();
      if (yr !== 2026) {
        log.corrections.push(`ID ${rec.ID}: Tarikh ${rec.TarikhDisplay} (tahun ${yr}) — bukan tahun 2026. Rekod dikekalkan untuk semakan.`);
      }
    }

    // Semak ID pendua
    if (idSeen[rec.ID]) {
      log.duplicates.push({ id: rec.ID, rec });
    } else {
      idSeen[rec.ID] = true;
    }

    // Semak status vs catatan
    const issues = detectQualityIssues(rec);
    if (issues.length > 0) {
      log.statusIssues.push({ rec, issues });
    }
  });

  return log;
}

/* =====================================================
   6. IMPORT FAIL EXCEL
   Format fail: Hari | Tarikh | Program/Aktiviti |
   Peserta | Jabatan/Unit | PIC | Catatan | Status
   ===================================================== */
function parseExcelFile(arrayBuffer, filename) {
  const wb = XLSX.read(arrayBuffer, { type:'array', cellDates:false, raw:false });
  let allRecords = [];
  let rawCount = 0;

  // Cuba semua helaian
  const sheets = wb.SheetNames;
  for (const sheetName of sheets) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    // Cuba baca dengan header (auto detect baris 1)
    let rows = XLSX.utils.sheet_to_json(ws, { defval:'', raw:false });
    if (!rows.length) {
      // Cuba tanpa header
      const aoa = XLSX.utils.sheet_to_json(ws, { defval:'', header:1, raw:false });
      if (aoa.length < 2) continue;
      // Cari baris header (mengandungi "tarikh" atau "aktiviti" atau "program")
      let headerRow = -1;
      for (let r = 0; r < Math.min(aoa.length, 10); r++) {
        const rowStr = aoa[r].join(' ').toLowerCase();
        if (rowStr.includes('tarikh') || rowStr.includes('aktiviti') || rowStr.includes('program')) {
          headerRow = r;
          break;
        }
      }
      if (headerRow < 0) continue;
      const headers = aoa[headerRow];
      rows = [];
      for (let r = headerRow+1; r < aoa.length; r++) {
        const rowData = aoa[r];
        if (rowData.every(v => !v || String(v).trim() === '')) continue;
        const obj = {};
        headers.forEach((h, ci) => { if (h) obj[h] = rowData[ci] || ''; });
        rows.push(obj);
      }
    }

    // Tentukan bulan dari nama fail atau helaian
    let bulanHint = detectMonthFromName(filename + ' ' + sheetName);

    rawCount += rows.length;

    rows.forEach((row, ri) => {
      // Skip baris kosong atau baris tajuk berulang
      const rowStr = Object.values(row).join(' ').toLowerCase();
      if (rowStr.includes('program / aktiviti') || rowStr.includes('program/aktiviti')) return;
      if (Object.values(row).every(v => !String(v||'').trim())) return;

      const namaVal    = findColValue(row, COLUMN_MAP.nama);
      const tarikhVal  = findColValue(row, COLUMN_MAP.tarikh);
      const hariVal    = findColValue(row, COLUMN_MAP.hari);
      const pesertaVal = findColValue(row, COLUMN_MAP.peserta);
      const jabatanVal = findColValue(row, COLUMN_MAP.jabatan);
      const picVal     = findColValue(row, COLUMN_MAP.pic);
      const catatanVal = findColValue(row, COLUMN_MAP.catatan);
      const statusVal  = findColValue(row, COLUMN_MAP.status);

      // Skip jika tiada nama aktiviti dan tiada tarikh
      if (!namaVal && !tarikhVal) return;
      // Skip baris tajuk dalam helaian
      if (normKey(namaVal) === 'program / aktiviti' || normKey(namaVal) === 'program/aktiviti') return;
      if (normKey(namaVal) === 'bil' || normKey(namaVal) === 'bil.') return;

      // Parse tarikh
      let tarikhISO = null;
      let tarikhDisplay = tarikhVal;
      const parsedDate = parseDateExcel(tarikhVal);
      if (parsedDate && !isNaN(parsedDate.getTime())) {
        tarikhISO = parsedDate.toISOString().slice(0,10);
        tarikhDisplay = fmtDateObj(parsedDate);
      }

      // Jika tiada tarikh, gunakan tarikh dari bulan hint
      if (!tarikhISO && bulanHint) {
        tarikhISO = `2026-${String(bulanHint).padStart(2,'0')}-01`;
        tarikhDisplay = `(Bulan ${MONTH_NAMES[bulanHint-1]} 2026 — tarikh tidak dikenal pasti)`;
      }

      const statusNorm = normaliseStatus(statusVal);
      const id = generateID(filename, sheetName, ri);

      const rec = {
        ID:            id,
        Sumber:        filename,
        Helaian:       sheetName,
        BulanHint:     bulanHint,
        Hari:          hariVal,
        TarikhRaw:     tarikhVal,
        TarikhISO:     tarikhISO,
        TarikhDisplay: tarikhDisplay,
        NamaAktiviti:  namaVal,
        Peserta:       pesertaVal,
        Jabatan:       normaliseJabatan(jabatanVal),
        JabatanAsal:   jabatanVal,
        PIC:           picVal,
        Catatan:       catatanVal,
        StatusAsal:    statusVal,
        StatusNormal:  statusNorm,
        IsuKualiti:    [],
      };
      allRecords.push(rec);
    });
  }

  return { records: allRecords, rawCount };
}

function generateID(filename, sheet, idx) {
  // Ekstrak prefix dari nama fail
  const fn = filename.replace(/\.xlsx?$/i,'').replace(/[^\w]/g,'').toUpperCase().substring(0,8);
  return `${fn}-${String(idx+1).padStart(3,'0')}`;
}

function detectMonthFromName(s) {
  const sl = s.toLowerCase();
  const months = [
    ['januari','january','jan'],
    ['februari','february','feb'],
    ['mac','march','mar'],
    ['april','apr'],
    ['mei','may'],
    ['jun','june'],
    ['julai','july','jul'],
    ['ogos','august','aug'],
    ['september','sep','sept'],
    ['oktober','october','oct','okt'],
    ['november','nov'],
    ['disember','december','dec','dis'],
  ];
  for (let i=0; i<months.length; i++) {
    if (months[i].some(m => sl.includes(m))) return i+1;
  }
  return null;
}

function normaliseJabatan(s) {
  if (!s) return '';
  let j = String(s).trim();
  // Standardise common variations
  j = j.replace(/\bja\b/gi, 'Jabatan');
  j = j.replace(/\bDPT\b/g, 'Jabatan');
  // Capitalize each word
  return j;
}

/* =====================================================
   7. MERGE DAN DEDUPLICATE
   ===================================================== */
function mergeAndDedup(newRecords) {
  // Deduplicate by exact match of NamaAktiviti+Jabatan+TarikhISO
  const seen = new Set();
  const deduped = [];
  const dupList = [];
  for (const rec of newRecords) {
    const key = `${normKey(rec.NamaAktiviti)}|${normKey(rec.Jabatan)}|${rec.TarikhISO||''}`;
    if (seen.has(key)) {
      dupList.push(rec);
    } else {
      seen.add(key);
      deduped.push(rec);
    }
  }
  return { deduped, dupList };
}

/* =====================================================
   8. PENGIRAAN STATISTIK
   ===================================================== */
function computeStats(data) {
  const total      = data.length;
  const selesai    = data.filter(a => a.StatusNormal === 'Selesai').length;
  const ditunda    = data.filter(a => a.StatusNormal === 'Ditunda').length;
  const digugurkan = data.filter(a => a.StatusNormal === 'Digugurkan').length;
  const lain       = total - selesai - ditunda - digugurkan;
  const pctSelesai    = calcPct(selesai, total);
  const pctDitunda    = calcPct(ditunda, total);
  const pctDigugurkan = calcPct(digugurkan, total);
  return { total, selesai, ditunda, digugurkan, lain, pctSelesai, pctDitunda, pctDigugurkan };
}

function computeMonthlyStats(data) {
  const stats = {};
  for (let m = 1; m <= 12; m++) {
    const d = data.filter(a => getMonth(a) === m);
    stats[m] = computeStats(d);
    stats[m].name = MONTH_NAMES[m-1];
    stats[m].hasData = d.length > 0;
  }
  return stats;
}

// Bulan yang mempunyai data dalam APP.data
function getMonthsWithData() {
  const set = new Set();
  APP.data.forEach(a => { const m = getMonth(a); if (m) set.add(m); });
  return [...set].sort((a,b)=>a-b);
}

// Suku yang mempunyai data
function getQuartersWithData() {
  const set = new Set();
  APP.data.forEach(a => { const q = getQuarter(getMonth(a)); if (q) set.add(q); });
  return [...set].sort();
}

function computeJabatanStats(data) {
  const jabMap = {};
  data.forEach(a => {
    const j = a.Jabatan || 'Tidak Direkodkan';
    if (!jabMap[j]) jabMap[j] = [];
    jabMap[j].push(a);
  });
  return Object.entries(jabMap).map(([j, recs]) => {
    const s = computeStats(recs);
    const pics = [...new Set(recs.map(r=>r.PIC).filter(Boolean))].join(', ');
    return { jabatan:j, pics, ...s, recs };
  }).sort((a,b) => b.total - a.total);
}

function getUniqueJabatan() {
  return [...new Set(APP.data.map(a=>a.Jabatan).filter(Boolean))].sort();
}

/* =====================================================
   9. RENDER KPI CARDS
   ===================================================== */
function renderKPIDashboard() {
  const s = computeStats(APP.data);
  const jabatan = getUniqueJabatan();
  const isuCount = APP.qualityLog.statusIssues.length;

  const cards = [
    { label:'Jumlah Aktiviti Direkodkan', value: s.total,        cls:'',              icon:'📋' },
    { label:'Aktiviti Selesai',           value: s.selesai,      cls:'kpi-selesai',   icon:'✅' },
    { label:'Aktiviti Ditunda',           value: s.ditunda,      cls:'kpi-ditunda',   icon:'⏸' },
    { label:'Aktiviti Digugurkan',        value: s.digugurkan,   cls:'kpi-isu',       icon:'🚫' },
    { label:'% Selesai (daripada rekod)', value: s.pctSelesai,   cls:'kpi-pct',       icon:'📊' },
    { label:'% Ditunda (daripada rekod)', value: s.pctDitunda,   cls:'kpi-ditunda',   icon:'📉' },
    { label:'Rekod Perlu Pengesahan',     value: isuCount,       cls:'kpi-isu',       icon:'⚠' },
    { label:'Jabatan/Unit Terlibat',      value: jabatan.length, cls:'kpi-jabatan',   icon:'🏢' },
  ];

  document.getElementById('kpiGrid').innerHTML = cards.map(c=>`
    <div class="kpi-card ${c.cls}">
      <div class="kpi-icon">${c.icon}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-label">${c.label}</div>
    </div>`).join('');
  document.getElementById('kpiNote').style.display = APP.hasData ? 'block' : 'none';
}

function renderKPISimple(containerId, data) {
  const s = computeStats(data);
  const cards = [
    { label:'Jumlah Rekod',  value:s.total,        cls:'',           icon:'📋' },
    { label:'Selesai',       value:s.selesai,      cls:'kpi-selesai',icon:'✅' },
    { label:'Ditunda',       value:s.ditunda,      cls:'kpi-ditunda',icon:'⏸' },
    { label:'Digugurkan',    value:s.digugurkan,   cls:'kpi-isu',    icon:'🚫' },
    { label:'% Selesai',     value:s.pctSelesai,   cls:'kpi-pct',    icon:'📊' },
  ];
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = cards.map(c=>`
    <div class="kpi-card ${c.cls}">
      <div class="kpi-icon">${c.icon}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-label">${c.label}</div>
    </div>`).join('');
}

/* =====================================================
   10. CHART HELPERS
   ===================================================== */
const CC = {
  selesai: '#16a34a', ditunda: '#b45309', digugurkan: '#dc2626', lain: '#6b7280',
  jan:'#1565c0', feb:'#7c3aed', mac:'#0891b2', line:'#1565c0',
  barBg: '#93c5fd',
};

function destroyChart(id) {
  if (APP.charts[id]) { APP.charts[id].destroy(); delete APP.charts[id]; }
}

function makeBarChart(id, labels, datasets, opts={}) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  APP.charts[id] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive:true, maintainAspectRatio:true,
      plugins:{ legend:{ position:'bottom' }, ...opts.plugins },
      scales:{ y:{ beginAtZero:true, ticks:{ precision:0 } }, ...opts.scales }
    }
  });
}

function makeHBarChart(id, labels, datasets) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  APP.charts[id] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:true,
      plugins:{ legend:{ position:'bottom' } },
      scales:{ x:{ beginAtZero:true, ticks:{ precision:0 } } }
    }
  });
}

function makeDonutChart(id, labels, values, colors) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  APP.charts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets:[{ data:values, backgroundColor:colors, borderWidth:2 }] },
    options: { responsive:true, maintainAspectRatio:true, plugins:{ legend:{ position:'bottom' } } }
  });
}

function makeLineChart(id, labels, datasets) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  APP.charts[id] = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive:true, maintainAspectRatio:true,
      plugins:{ legend:{ position:'bottom' } },
      scales:{ y:{ beginAtZero:true, ticks:{ precision:0 } } }
    }
  });
}

/* =====================================================
   11. RENDER DASHBOARD UTAMA
   ===================================================== */
function renderDashboard() {
  document.getElementById('importNotice').style.display = APP.hasData ? 'none' : 'flex';

  if (!APP.hasData) {
    renderKPIDashboard();
    return;
  }

  renderKPIDashboard();

  const monthly = computeMonthlyStats(APP.data);
  const activeMonths = getMonthsWithData();
  const monthLabels = activeMonths.map(m => MONTH_NAMES[m-1].substring(0,3));
  const ms = activeMonths.map(m => monthly[m]);

  // Bar: Selesai vs Ditunda vs Digugurkan per bulan (bulan yang ada data sahaja)
  makeBarChart('chartBarMain', monthLabels, [
    { label:'Selesai',    data:ms.map(s=>s.selesai),    backgroundColor:CC.selesai },
    { label:'Ditunda',    data:ms.map(s=>s.ditunda),    backgroundColor:CC.ditunda },
    { label:'Digugurkan', data:ms.map(s=>s.digugurkan), backgroundColor:CC.digugurkan },
  ]);

  // Donut: status keseluruhan
  const st = computeStats(APP.data);
  makeDonutChart('chartDonutMain',
    ['Selesai','Ditunda','Digugurkan'],
    [st.selesai, st.ditunda, st.digugurkan],
    [CC.selesai, CC.ditunda, CC.digugurkan]
  );

  // Line: trend bilangan aktiviti (bulan yang ada data)
  makeLineChart('chartLineMain', monthLabels, [
    { label:'Jumlah Rekod', data:ms.map(s=>s.total),
      borderColor:CC.line, backgroundColor:'rgba(21,101,192,0.1)',
      fill:true, tension:0.3, pointRadius:5 },
    { label:'Selesai', data:ms.map(s=>s.selesai),
      borderColor:CC.selesai, backgroundColor:'rgba(22,163,74,0.08)',
      fill:false, tension:0.3, pointRadius:5 },
  ]);

  // HBar: jabatan
  const jabs = computeJabatanStats(APP.data).slice(0, 12);
  makeHBarChart('chartBarJabatan',
    jabs.map(j => j.jabatan.replace('Jabatan ','').replace('Unit ','')),
    [
      { label:'Selesai',    data:jabs.map(j=>j.selesai),    backgroundColor:CC.selesai },
      { label:'Ditunda',    data:jabs.map(j=>j.ditunda),    backgroundColor:CC.ditunda },
      { label:'Digugurkan', data:jabs.map(j=>j.digugurkan), backgroundColor:CC.digugurkan },
    ]
  );

  // Table terkini
  const recent = [...APP.data].slice(-10).reverse();
  document.getElementById('tbodyRecent').innerHTML = recent.map(a=>`
    <tr>
      <td style="font-size:11px;color:#888">${a.ID}</td>
      <td>${a.NamaAktiviti||'—'}</td>
      <td>${a.Jabatan||'—'}</td>
      <td>${a.Peserta||'—'}</td>
      <td>${a.PIC||'—'}</td>
      <td>${badgeHtml(a.StatusNormal)}</td>
      <td>${a.TarikhDisplay||'—'}</td>
    </tr>`).join('') || '<tr><td colspan="7" style="text-align:center;color:#888;padding:20px">Tiada rekod</td></tr>';
}

/* =====================================================
   12. RENDER LAPORAN BULANAN
   ===================================================== */
function renderBulanan() {
  const sel = document.getElementById('filterBulan').value;
  const m = (sel !== 'all') ? parseInt(sel) : null;

  const unavail = document.getElementById('bulananUnavailable');
  const activeMonths = getMonthsWithData();

  // Jika bulan dipilih tapi tiada data
  if (m && !activeMonths.includes(m)) {
    unavail.style.display = 'block';
    document.getElementById('kpiBulanan').innerHTML = '';
    document.getElementById('tbodyBulananDitunda').innerHTML = '';
    document.getElementById('tbodyBulanan').innerHTML = '';
    destroyChart('chartBulananBar'); destroyChart('chartBulananDonut');
    return;
  }
  unavail.style.display = 'none';

  const data = m ? APP.data.filter(a => getMonth(a) === m) : APP.data;
  renderKPISimple('kpiBulanan', data);

  // Bar: selesai vs ditunda vs digugurkan
  if (sel === 'all') {
    const monthly = computeMonthlyStats(APP.data);
    const mLabels = activeMonths.map(mn => MONTH_NAMES[mn-1].substring(0,3));
    makeBarChart('chartBulananBar', mLabels, [
      { label:'Selesai',    data:activeMonths.map(mn=>monthly[mn].selesai),    backgroundColor:CC.selesai },
      { label:'Ditunda',    data:activeMonths.map(mn=>monthly[mn].ditunda),    backgroundColor:CC.ditunda },
      { label:'Digugurkan', data:activeMonths.map(mn=>monthly[mn].digugurkan), backgroundColor:CC.digugurkan },
    ]);
  } else {
    // Bar mengikut jabatan dalam bulan
    const jabs = computeJabatanStats(data);
    makeBarChart('chartBulananBar',
      jabs.map(j=>j.jabatan.replace('Jabatan ','').replace('Unit ','').substring(0,20)),
      [
        { label:'Selesai',    data:jabs.map(j=>j.selesai),    backgroundColor:CC.selesai },
        { label:'Ditunda',    data:jabs.map(j=>j.ditunda),    backgroundColor:CC.ditunda },
        { label:'Digugurkan', data:jabs.map(j=>j.digugurkan), backgroundColor:CC.digugurkan },
      ]
    );
  }

  const s = computeStats(data);
  makeDonutChart('chartBulananDonut',
    ['Selesai','Ditunda','Digugurkan'],
    [s.selesai, s.ditunda, s.digugurkan],
    [CC.selesai, CC.ditunda, CC.digugurkan]
  );

  // Senarai ditunda
  const ditunda = data.filter(a => a.StatusNormal === 'Ditunda');
  document.getElementById('tbodyBulananDitunda').innerHTML = ditunda.length ?
    ditunda.map((a,i)=>`<tr>
      <td>${i+1}</td>
      <td>${a.NamaAktiviti||'—'}</td>
      <td>${a.Jabatan||'—'}</td>
      <td>${a.TarikhDisplay||'—'}</td>
      <td>${a.PIC||'—'}</td>
      <td>${badgeHtml(a.StatusNormal)}</td>
      <td style="max-width:220px;font-size:12px">${a.Catatan||'—'}</td>
    </tr>`).join('') :
    '<tr><td colspan="7" style="text-align:center;color:#888;padding:16px">Tiada aktiviti ditunda dalam tempoh ini.</td></tr>';

  // Semua aktiviti
  document.getElementById('tbodyBulanan').innerHTML = data.map(a=>`
    <tr>
      <td style="font-size:11px;color:#888">${a.ID}</td>
      <td>${a.NamaAktiviti||'—'}</td>
      <td>${a.Jabatan||'—'}</td>
      <td>${a.TarikhDisplay||'—'}</td>
      <td>${a.Peserta||'—'}</td>
      <td>${a.PIC||'—'}</td>
      <td>${badgeHtml(a.StatusNormal)}</td>
      <td style="max-width:180px;font-size:12px">${a.Catatan||'—'}</td>
    </tr>`).join('') || '<tr><td colspan="8" style="text-align:center;color:#888;padding:16px">Tiada rekod</td></tr>';
}

/* =====================================================
   13. RENDER LAPORAN SUKU TAHUNAN
   ===================================================== */
function renderSuku() {
  const sel = document.getElementById('filterSuku').value;
  const activeMonths = getMonthsWithData();
  const qWithData = getQuartersWithData();

  // Data untuk suku dipilih atau semua suku yang ada data
  let dispData = APP.data;
  let dispMonths = activeMonths;
  if (sel !== 'all') {
    const qMonthMap = { Q1:[1,2,3], Q2:[4,5,6], Q3:[7,8,9], Q4:[10,11,12] };
    dispData = APP.data.filter(a => getQuarter(getMonth(a)) === sel);
    dispMonths = activeMonths.filter(m => qMonthMap[sel].includes(m));
  }

  renderKPISimple('kpiSuku', dispData);

  const monthly = computeMonthlyStats(APP.data);
  const mLabels = dispMonths.map(m => MONTH_NAMES[m-1].substring(0,3)+' 2026');

  makeBarChart('chartSukuBar', mLabels, [
    { label:'Selesai',    data:dispMonths.map(m=>monthly[m].selesai),    backgroundColor:CC.selesai },
    { label:'Ditunda',    data:dispMonths.map(m=>monthly[m].ditunda),    backgroundColor:CC.ditunda },
    { label:'Digugurkan', data:dispMonths.map(m=>monthly[m].digugurkan), backgroundColor:CC.digugurkan },
  ]);

  const st = computeStats(dispData);
  makeDonutChart('chartSukuDonut',
    ['Selesai','Ditunda','Digugurkan'],
    [st.selesai, st.ditunda, st.digugurkan],
    [CC.selesai, CC.ditunda, CC.digugurkan]
  );

  // Jadual mengikut bulan yang ada data
  document.getElementById('tbodySuku').innerHTML = dispMonths.map(m => {
    const s = monthly[m];
    return `<tr>
      <td><strong>${MONTH_NAMES[m-1]} 2026</strong></td>
      <td class="tnum">${s.total}</td>
      <td class="tnum">${s.selesai}</td>
      <td class="tnum">${s.ditunda}</td>
      <td class="tnum">${s.digugurkan}</td>
      <td><strong>${s.pctSelesai}</strong></td>
      <td><strong>${s.pctDitunda}</strong></td>
    </tr>`;
  }).join('') || '<tr><td colspan="7" style="text-align:center;color:#888;padding:16px">Tiada data untuk tempoh ini.</td></tr>';

  // Kemaskini jadual "Data Belum Dibekalkan" — hanya suku yang tiada data
  const allQ = ['Q1','Q2','Q3','Q4'];
  const qLabels = { Q1:'Januari – Mac', Q2:'April – Jun', Q3:'Julai – September', Q4:'Oktober – Disember' };
  const unavailRows = allQ.filter(q => !qWithData.includes(q)).map(q =>
    `<tr><td><strong>${q}</strong></td><td>${qLabels[q]} 2026</td><td><span class="badge badge-belum">Data Belum Dibekalkan</span></td></tr>`
  ).join('');
  const unavailTbody = document.querySelector('#page-suku table:last-of-type tbody');
  if (unavailTbody) unavailTbody.innerHTML = unavailRows ||
    '<tr><td colspan="3" style="text-align:center;color:#888;padding:12px">Semua suku tahunan telah mempunyai data.</td></tr>';
}

/* =====================================================
   14. RENDER ANALISIS JABATAN
   ===================================================== */
function renderJabatan() {
  let data = APP.data;
  const jabF = document.getElementById('filterJabatanPage').value;
  const blnF = document.getElementById('filterJabatanBulan').value;
  if (jabF !== 'all') data = data.filter(a=>a.Jabatan===jabF);
  if (blnF !== 'all') data = data.filter(a=>getMonth(a)===parseInt(blnF));

  renderKPISimple('kpiJabatan', data);

  const jabs = computeJabatanStats(data);
  const jabLabels = jabs.map(j=>j.jabatan.replace('Jabatan ','').replace('Unit ','').substring(0,25));

  makeHBarChart('chartJabatanBar', jabLabels, [
    { label:'Selesai',    data:jabs.map(j=>j.selesai),    backgroundColor:CC.selesai },
    { label:'Ditunda',    data:jabs.map(j=>j.ditunda),    backgroundColor:CC.ditunda },
    { label:'Digugurkan', data:jabs.map(j=>j.digugurkan), backgroundColor:CC.digugurkan },
  ]);

  makeHBarChart('chartJabatanPct', jabLabels, [{
    label:'% Selesai',
    data: jabs.map(j=>calcPctNum(j.selesai, j.total)),
    backgroundColor: 'rgba(21,101,192,0.75)',
  }]);

  document.getElementById('tbodyJabatan').innerHTML = jabs.map((j,i)=>`
    <tr>
      <td>${i+1}</td>
      <td>${j.jabatan}</td>
      <td class="tnum">${j.total}</td>
      <td class="tnum">${j.selesai}</td>
      <td class="tnum">${j.ditunda}</td>
      <td class="tnum">${j.digugurkan}</td>
      <td><strong>${j.pctSelesai}</strong></td>
      <td style="font-size:12px;max-width:150px">${j.pics||'—'}</td>
    </tr>`).join('') || '<tr><td colspan="8" style="text-align:center;color:#888;padding:16px">Tiada rekod</td></tr>';
}

/* =====================================================
   15. RENDER AKTIVITI DITUNDA
   ===================================================== */
function renderDitunda() {
  let data = APP.data.filter(a => a.StatusNormal === 'Ditunda');
  const jabF = document.getElementById('filterDitundaJabatan').value;
  const blnF = document.getElementById('filterDitundaBulan').value;
  if (jabF !== 'all') data = data.filter(a=>a.Jabatan===jabF);
  if (blnF !== 'all') data = data.filter(a=>getMonth(a)===parseInt(blnF));

  const kpiEl = document.getElementById('kpiDitunda');
  const s = computeStats(data);
  kpiEl.innerHTML = `
    <div class="kpi-card kpi-ditunda"><div class="kpi-icon">⏸</div><div class="kpi-value">${data.length}</div><div class="kpi-label">Jumlah Ditunda</div></div>
    <div class="kpi-card"><div class="kpi-icon">🏢</div><div class="kpi-value">${[...new Set(data.map(a=>a.Jabatan))].length}</div><div class="kpi-label">Jabatan Terlibat</div></div>
    <div class="kpi-card"><div class="kpi-icon">📋</div><div class="kpi-value">${APP.data.length}</div><div class="kpi-label">Jumlah Semua Rekod</div></div>
    <div class="kpi-card kpi-pct"><div class="kpi-icon">📊</div><div class="kpi-value">${calcPct(APP.data.filter(a=>a.StatusNormal==='Ditunda').length, APP.data.length)}</div><div class="kpi-label">% Ditunda Keseluruhan</div></div>
  `;

  document.getElementById('ditundaTableTitle').textContent =
    `Senarai Aktiviti Ditunda (${data.length} rekod)`;

  document.getElementById('tbodyDitunda').innerHTML = data.length ? data.map((a,i)=>{
    // Cadangkan tindakan susulan berdasarkan catatan
    let tindakan = 'Dapatkan pengesahan status terkini daripada jabatan berkenaan.';
    const cat = String(a.Catatan||'').toLowerCase();
    if (cat.includes('april') || cat.includes('mei') || cat.includes('jun')) {
      tindakan = 'Semak jadual baharu yang dicadangkan dalam catatan.';
    } else if (cat.includes('penceramah') || cat.includes('fasilitator')) {
      tindakan = 'Dapatkan penceramah/fasilitator alternatif atau tentukan tarikh baharu.';
    } else if (cat.includes('pertindihan') || cat.includes('konflik')) {
      tindakan = 'Selesaikan pertindihan jadual dan tetapkan tarikh baharu.';
    } else if (cat.includes('dokumen') || cat.includes('prosedur')) {
      tindakan = 'Selesaikan isu dokumen/prosedur dan jadualkan semula.';
    }
    return `<tr>
      <td>${i+1}</td>
      <td>${a.NamaAktiviti||'—'}</td>
      <td>${a.Jabatan||'—'}</td>
      <td>${a.TarikhDisplay||'—'}</td>
      <td>${a.PIC||'—'}</td>
      <td style="max-width:220px;font-size:12px">${a.Catatan||'—'}</td>
      <td style="font-size:12px;color:#1565c0">${tindakan}</td>
    </tr>`;
  }).join('') :
  '<tr><td colspan="7" style="text-align:center;color:#888;padding:20px">Tiada aktiviti ditunda dalam pilihan semasa.</td></tr>';
}

/* =====================================================
   16. RENDER LOG KUALITI DATA
   ===================================================== */
function renderKualiti() {
  const ql = APP.qualityLog;
  const kpiEl = document.getElementById('kpiKualiti');
  kpiEl.innerHTML = `
    <div class="kpi-card"><div class="kpi-icon">📥</div><div class="kpi-value">${ql.totalRaw}</div><div class="kpi-label">Rekod Asal (Sebelum Pembersihan)</div></div>
    <div class="kpi-card kpi-selesai"><div class="kpi-icon">✅</div><div class="kpi-value">${ql.totalAfterMerge}</div><div class="kpi-label">Rekod Selepas Gabungan</div></div>
    <div class="kpi-card kpi-ditunda"><div class="kpi-icon">📭</div><div class="kpi-value">${ql.emptyFields.length}</div><div class="kpi-label">Rekod dengan Medan Kosong</div></div>
    <div class="kpi-card kpi-isu"><div class="kpi-icon">⚠</div><div class="kpi-value">${ql.statusIssues.length}</div><div class="kpi-label">Rekod Percanggahan Status</div></div>
    <div class="kpi-card"><div class="kpi-icon">🔁</div><div class="kpi-value">${ql.duplicates.length}</div><div class="kpi-label">Rekod Pendua Disyaki</div></div>
    <div class="kpi-card"><div class="kpi-icon">📅</div><div class="kpi-value">${ql.dateIssues.length}</div><div class="kpi-label">Rekod Tarikh Meragukan</div></div>
  `;

  // Log ringkasan
  const logRows = [
    ['Jumlah rekod asal (sebelum pembersihan)', ql.totalRaw, 'Rekod mentah dari semua fail Excel yang diimport'],
    ['Jumlah rekod selepas penggabungan', ql.totalAfterMerge, 'Rekod selepas penggabungan dan pembersihan asas'],
    ['Rekod baris kosong/tajuk berulang dibuang', ql.totalRaw - ql.totalAfterMerge - ql.duplicates.length, 'Baris tanpa nama aktiviti dan tarikh ditolak'],
    ['Rekod pendua disyaki', ql.duplicates.length, 'Berdasarkan padanan Nama Aktiviti + Jabatan + Tarikh'],
    ['Rekod dengan medan kosong', ql.emptyFields.length, 'Medan yang tiada: Program/Aktiviti, Jabatan, Tarikh, Status, atau PIC'],
    ['Rekod dengan tarikh meragukan', ql.dateIssues.length, 'Nilai tarikh tidak dapat dikenali formatnya'],
    ['Rekod dengan percanggahan status-catatan', ql.statusIssues.length, 'Status rekod tidak konsisten dengan kandungan catatan'],
    ['Pembetulan status yang dibuat', ql.corrections.filter(c=>c.includes('Status')).length, 'Status SELESAI/DITUNDA/DIGUGURKAN dinormalkan kepada format piawai'],
  ];

  document.getElementById('tbodyKualitiLog').innerHTML = logRows.map(r=>`
    <tr>
      <td>${r[0]}</td>
      <td class="tnum"><strong>${r[1]}</strong></td>
      <td style="font-size:12px;color:#57606a">${r[2]}</td>
    </tr>`).join('');

  // Isu status
  document.getElementById('tbodyKualitiIsu').innerHTML = ql.statusIssues.length ?
    ql.statusIssues.map(({rec, issues})=>`
      <tr>
        <td style="font-size:11px">${rec.ID}</td>
        <td>${rec.NamaAktiviti||'—'}</td>
        <td>${rec.Jabatan||'—'}</td>
        <td>${badgeHtml(rec.StatusNormal)}</td>
        <td style="max-width:200px;font-size:12px">${rec.Catatan||'—'}</td>
        <td class="badge-isu" style="font-size:12px;color:#dc2626">${issues.join('; ')}</td>
      </tr>`).join('') :
    '<tr><td colspan="6" style="text-align:center;color:#888;padding:16px">Tiada percanggahan status-catatan dikesan.</td></tr>';

  // Medan kosong
  document.getElementById('tbodyKualitiKosong').innerHTML = ql.emptyFields.length ?
    ql.emptyFields.slice(0,50).map(({rec, empty})=>`
      <tr>
        <td style="font-size:11px">${rec.ID}</td>
        <td>${rec.NamaAktiviti||'(tiada nama)'}</td>
        <td>${rec.Jabatan||'—'}</td>
        <td style="color:#b45309;font-size:12px">${empty.join(', ')}</td>
        <td>${badgeHtml(rec.StatusNormal)}</td>
      </tr>`).join('') :
    '<tr><td colspan="5" style="text-align:center;color:#888;padding:16px">Tiada rekod dengan medan penting kosong.</td></tr>';
}

/* =====================================================
   17. RENDER SENARAI AKTIVITI
   ===================================================== */
function filterSenaraiData() {
  const cari  = (document.getElementById('cariAktiviti').value||'').toLowerCase();
  const jab   = document.getElementById('filterSenaraiJabatan').value;
  const stat  = document.getElementById('filterSenaraiStatus').value;
  const suku  = document.getElementById('filterSenaraiSuku').value;
  const urut  = document.getElementById('filterSenaraiUrutan').value;

  let data = APP.data.filter(a => {
    if (cari && !String(a.NamaAktiviti||'').toLowerCase().includes(cari)) return false;
    if (jab  !== 'all' && a.Jabatan !== jab) return false;
    if (stat !== 'all' && a.StatusNormal !== stat) return false;
    if (suku !== 'all' && getQuarter(getMonth(a)) !== suku) return false;
    return true;
  });

  data.sort((a,b) => {
    if (urut === 'nama')    return String(a.NamaAktiviti||'').localeCompare(String(b.NamaAktiviti||''),'ms');
    if (urut === 'jabatan') return String(a.Jabatan||'').localeCompare(String(b.Jabatan||''),'ms');
    const da = a.TarikhISO ? new Date(a.TarikhISO) : new Date(0);
    const db = b.TarikhISO ? new Date(b.TarikhISO) : new Date(0);
    return da - db;
  });
  return data;
}

function renderSenarai() {
  const all   = filterSenaraiData();
  const total = all.length;
  const pages = Math.ceil(total / APP.pageSize) || 1;
  if (APP.currentPage > pages) APP.currentPage = pages;

  const start = (APP.currentPage - 1) * APP.pageSize;
  const slice = all.slice(start, start + APP.pageSize);

  document.getElementById('tableInfo').textContent =
    `Menunjukkan ${total ? start+1 : 0}–${Math.min(start+APP.pageSize, total)} daripada ${total} rekod`;

  document.getElementById('tbodyAktiviti').innerHTML = slice.map(a=>`
    <tr>
      <td style="font-size:11px;color:#888">${a.ID}</td>
      <td>${a.Hari||'—'}</td>
      <td>${a.TarikhDisplay||'—'}</td>
      <td>${a.NamaAktiviti||'—'}</td>
      <td>${a.Peserta||'—'}</td>
      <td>${a.Jabatan||'—'}</td>
      <td>${a.PIC||'—'}</td>
      <td>${badgeHtml(a.StatusNormal)}</td>
      <td style="max-width:140px;font-size:12px">${a.Catatan||'—'}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="openModal('${a.ID}')">Lihat</button>
        <button class="btn btn-sm btn-primary" onclick="editAktiviti('${a.ID}')">Edit</button>
        <button class="btn btn-sm btn-danger" onclick="deleteAktiviti('${a.ID}')">Padam</button>
      </td>
    </tr>`).join('') || '<tr><td colspan="10" style="text-align:center;color:#888;padding:24px">Tiada rekod. Sila import data Excel terlebih dahulu.</td></tr>';

  // Pagination
  const pg = document.getElementById('pagination');
  if (pages <= 1) { pg.innerHTML = ''; return; }
  let html = `<button class="page-btn" onclick="goPage(${APP.currentPage-1})" ${APP.currentPage===1?'disabled':''}>&#8249;</button>`;
  const startPg = Math.max(1, APP.currentPage-2);
  const endPg   = Math.min(pages, APP.currentPage+2);
  if (startPg > 1) html += `<button class="page-btn" onclick="goPage(1)">1</button><span style="padding:0 4px">…</span>`;
  for (let i=startPg; i<=endPg; i++) {
    html += `<button class="page-btn ${i===APP.currentPage?'active':''}" onclick="goPage(${i})">${i}</button>`;
  }
  if (endPg < pages) html += `<span style="padding:0 4px">…</span><button class="page-btn" onclick="goPage(${pages})">${pages}</button>`;
  html += `<button class="page-btn" onclick="goPage(${APP.currentPage+1})" ${APP.currentPage===pages?'disabled':''}>&#8250;</button>`;
  pg.innerHTML = html;
}

function goPage(p) { APP.currentPage = p; renderSenarai(); }

/* =====================================================
   18. MODAL BUTIRAN
   ===================================================== */
function openModal(id) {
  const a = APP.data.find(x=>x.ID===id);
  if (!a) return;
  document.getElementById('modalTitle').textContent = a.NamaAktiviti || 'Butiran Rekod';
  const isu = a.IsuKualiti && a.IsuKualiti.length
    ? `<div class="detail-item detail-full"><span class="detail-label">Isu Kualiti Data</span><span class="detail-value" style="color:#dc2626">${a.IsuKualiti.join('; ')}</span></div>`
    : '';
  document.getElementById('modalBody').innerHTML = `
    <div class="detail-grid">
      <div class="detail-item"><span class="detail-label">ID Rekod</span><span class="detail-value">${a.ID}</span></div>
      <div class="detail-item"><span class="detail-label">Sumber Fail</span><span class="detail-value">${a.Sumber}</span></div>
      <div class="detail-item"><span class="detail-label">Hari</span><span class="detail-value">${a.Hari||'—'}</span></div>
      <div class="detail-item"><span class="detail-label">Tarikh</span><span class="detail-value">${a.TarikhDisplay||'—'}</span></div>
      <div class="detail-item detail-full"><span class="detail-label">Program / Aktiviti</span><span class="detail-value">${a.NamaAktiviti||'—'}</span></div>
      <div class="detail-item"><span class="detail-label">Jabatan / Unit</span><span class="detail-value">${a.Jabatan||'—'}</span></div>
      <div class="detail-item"><span class="detail-label">Peserta</span><span class="detail-value">${a.Peserta||'—'}</span></div>
      <div class="detail-item"><span class="detail-label">Pegawai Urusetia (PIC)</span><span class="detail-value">${a.PIC||'—'}</span></div>
      <div class="detail-item"><span class="detail-label">Status Rekod (Asal)</span><span class="detail-value">${a.StatusAsal||'—'}</span></div>
      <div class="detail-item"><span class="detail-label">Status (Dinormalkan)</span><span class="detail-value">${badgeHtml(a.StatusNormal)}</span></div>
      <div class="detail-item detail-full"><span class="detail-label">Catatan / Justifikasi</span><span class="detail-value">${a.Catatan||'—'}</span></div>
      ${isu}
    </div>`;
  document.getElementById('modalOverlay').classList.add('open');
}

document.getElementById('modalClose').addEventListener('click', ()=>{
  document.getElementById('modalOverlay').classList.remove('open');
});
document.getElementById('modalOverlay').addEventListener('click', e=>{
  if (e.target===document.getElementById('modalOverlay'))
    document.getElementById('modalOverlay').classList.remove('open');
});

/* =====================================================
   19. BORANG TAMBAH/EDIT
   ===================================================== */
function clearForm() {
  document.getElementById('editId').value = '';
  document.getElementById('formAktiviti').reset();
  document.getElementById('btnSimpan').textContent = '✓ Simpan Rekod';
}

function editAktiviti(id) {
  const a = APP.data.find(x=>x.ID===id);
  if (!a) return;
  navigateTo('pengurusan');
  document.getElementById('editId').value = id;
  document.getElementById('fIDAktiviti').value  = a.ID;
  document.getElementById('fHari').value        = a.Hari||'';
  document.getElementById('fTarikh').value      = a.TarikhISO||'';
  document.getElementById('fNamaAktiviti').value= a.NamaAktiviti||'';
  document.getElementById('fJabatan').value     = a.Jabatan||'';
  document.getElementById('fPeserta').value     = a.Peserta||'';
  document.getElementById('fPegawai').value     = a.PIC||'';
  document.getElementById('fStatus').value      = a.StatusNormal||'';
  document.getElementById('fCatatan').value     = a.Catatan||'';
  document.getElementById('btnSimpan').textContent = '✓ Kemas Kini Rekod';
  document.getElementById('fIDAktiviti').scrollIntoView({ behavior:'smooth' });
}

function deleteAktiviti(id) {
  if (!confirm(`Padam rekod "${id}"? Tindakan ini tidak boleh dibuat asal.`)) return;
  APP.data = APP.data.filter(x=>x.ID!==id);
  saveToStorage();
  refreshAll();
  showToast('Rekod berjaya dipadam.', 'success');
}

document.getElementById('formAktiviti').addEventListener('submit', e => {
  e.preventDefault();
  const editId = document.getElementById('editId').value;
  const newId  = document.getElementById('fIDAktiviti').value.trim();
  if (!editId && APP.data.find(x=>x.ID===newId)) {
    showToast('ID rekod sudah wujud.', 'error'); return;
  }
  const tarikhVal = document.getElementById('fTarikh').value;
  const d = tarikhVal ? new Date(tarikhVal) : null;
  const rec = {
    ID:          newId,
    Sumber:      'Manual',
    Helaian:     'Manual',
    BulanHint:   d ? d.getMonth()+1 : null,
    Hari:        document.getElementById('fHari').value.trim(),
    TarikhRaw:   tarikhVal,
    TarikhISO:   tarikhVal||null,
    TarikhDisplay: d ? fmtDateObj(d) : (tarikhVal||'—'),
    NamaAktiviti: document.getElementById('fNamaAktiviti').value.trim(),
    Peserta:     document.getElementById('fPeserta').value.trim(),
    Jabatan:     document.getElementById('fJabatan').value.trim(),
    JabatanAsal: document.getElementById('fJabatan').value.trim(),
    PIC:         document.getElementById('fPegawai').value.trim(),
    StatusAsal:  document.getElementById('fStatus').value,
    StatusNormal:document.getElementById('fStatus').value,
    Catatan:     document.getElementById('fCatatan').value.trim(),
    IsuKualiti:  [],
  };

  if (editId) {
    const idx = APP.data.findIndex(x=>x.ID===editId);
    if (idx!==-1) APP.data[idx] = rec;
    showToast('Rekod berjaya dikemas kini.', 'success');
  } else {
    APP.data.push(rec);
    showToast('Rekod baharu berjaya ditambah.', 'success');
  }
  saveToStorage(); refreshAll(); clearForm();
});

document.getElementById('btnBatal').addEventListener('click', clearForm);

/* =====================================================
   20. LOCAL STORAGE
   ===================================================== */
const STORAGE_KEY = 'pbu2026_data';
const STORAGE_TS  = 'pbu2026_ts';
const STORAGE_QL  = 'pbu2026_ql';
const STORAGE_CNT = 'pbu2026_cnt';

function saveToStorage() {
  const ts = new Date().toISOString();
  APP.lastUpdated = ts;
  APP.hasData = APP.data.length > 0;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(APP.data));
    localStorage.setItem(STORAGE_TS,  ts);
    localStorage.setItem(STORAGE_QL,  JSON.stringify(APP.qualityLog));
    localStorage.setItem(STORAGE_CNT, String(APP.filesImported));
  } catch(e) {
    showToast('Gagal simpan ke LocalStorage. Data mungkin terlalu besar.', 'error');
  }
  updateStatusPanel();
}

function loadFromStorage() {
  const raw = localStorage.getItem(STORAGE_KEY);
  const ts  = localStorage.getItem(STORAGE_TS);
  const ql  = localStorage.getItem(STORAGE_QL);
  const cnt = localStorage.getItem(STORAGE_CNT);
  if (raw) {
    try {
      APP.data = JSON.parse(raw);
      APP.hasData = APP.data.length > 0;
      APP.lastUpdated = ts;
      APP.filesImported = parseInt(cnt)||0;
      if (ql) APP.qualityLog = JSON.parse(ql);
      return true;
    } catch(e) { return false; }
  }
  return false;
}

function updateStatusPanel() {
  const el = id => document.getElementById(id);
  el('statRekod').textContent = APP.data.length;
  el('statFail').textContent  = APP.filesImported;
  el('statKemas').textContent = APP.lastUpdated
    ? new Date(APP.lastUpdated).toLocaleString('ms-MY') : '—';
  el('statIsu').textContent   = APP.qualityLog.statusIssues
    ? APP.qualityLog.statusIssues.length : 0;

  el('lastUpdated').textContent = APP.lastUpdated
    ? new Date(APP.lastUpdated).toLocaleDateString('ms-MY') : '—';

  const badge = el('dataBadge');
  if (APP.hasData) {
    badge.textContent = `${APP.data.length} Rekod`;
    badge.className   = 'header-badge ready';
  } else {
    badge.textContent = 'TIADA DATA';
    badge.className   = 'header-badge';
  }
}

/* =====================================================
   21. IMPORT EXCEL — MULTI-FILE
   ===================================================== */
async function processFiles(files) {
  const logEl = document.getElementById('importLog');
  logEl.style.display = 'block';
  logEl.innerHTML = '<div style="font-weight:700;color:#0d2a4e;margin-bottom:8px">Log Import:</div>';

  let allNewRecords = [];
  let totalRaw = 0;

  for (const file of files) {
    logEl.innerHTML += `<div class="import-log-item"><span class="import-log-ok">→</span><span>Membaca fail: <strong>${file.name}</strong>...</span></div>`;

    try {
      const ab = await readFileAsArrayBuffer(file);
      const { records, rawCount } = parseExcelFile(ab, file.name);
      totalRaw += rawCount;
      allNewRecords = allNewRecords.concat(records);
      logEl.innerHTML += `<div class="import-log-item"><span class="import-log-ok">✓</span><span>${file.name}: ${rawCount} baris dibaca, ${records.length} rekod diproses.</span></div>`;
    } catch(err) {
      logEl.innerHTML += `<div class="import-log-item"><span class="import-log-err">✗</span><span>Ralat: ${file.name} — ${err.message}</span></div>`;
    }
  }

  if (!allNewRecords.length) {
    logEl.innerHTML += '<div class="import-log-item"><span class="import-log-err">✗</span><span>Tiada rekod dapat diproses. Semak format lajur dalam fail Excel.</span></div>';
    showToast('Tiada rekod diproses. Semak format fail.', 'error');
    return;
  }

  // Deduplicate
  const { deduped, dupList } = mergeAndDedup(allNewRecords);

  // Gabungkan dengan data sedia ada
  // (import baru menggantikan data sedia ada dari fail yang sama)
  const mergedSources = new Set(allNewRecords.map(r=>r.Sumber));
  const existing = APP.data.filter(r => !mergedSources.has(r.Sumber) && r.Sumber !== 'Manual');
  const finalData = [...existing, ...deduped];

  // Quality check
  const ql = runDataQualityCheck(finalData);
  ql.totalRaw = totalRaw + existing.length;

  // Tag isu pada rekod
  finalData.forEach(rec => {
    rec.IsuKualiti = detectQualityIssues(rec);
  });

  APP.data = finalData;
  APP.qualityLog = ql;
  APP.filesImported += files.length;
  APP.hasData = true;

  // Log
  logEl.innerHTML += `<div class="import-log-item"><span class="import-log-ok">✓</span><span>Jumlah rekod selepas gabungan: <strong>${finalData.length}</strong></span></div>`;
  if (dupList.length) {
    logEl.innerHTML += `<div class="import-log-item"><span class="import-log-warn">⚠</span><span>${dupList.length} rekod pendua disyaki (berdasarkan Nama+Jabatan+Tarikh) — dikekalkan satu sahaja.</span></div>`;
  }
  if (ql.statusIssues.length) {
    logEl.innerHTML += `<div class="import-log-item"><span class="import-log-warn">⚠</span><span>${ql.statusIssues.length} rekod dengan percanggahan status-catatan. Semak halaman Log Kualiti Data.</span></div>`;
  }
  if (ql.emptyFields.length) {
    logEl.innerHTML += `<div class="import-log-item"><span class="import-log-warn">⚠</span><span>${ql.emptyFields.length} rekod dengan medan penting kosong. Semak halaman Log Kualiti Data.</span></div>`;
  }
  logEl.innerHTML += `<div class="import-log-item"><span class="import-log-ok">✓</span><span style="font-weight:700">Import selesai. ${finalData.length} rekod dimuatkan ke dashboard.</span></div>`;

  saveToStorage();
  populateFilters();
  refreshAll();
  showToast(`${finalData.length} rekod berjaya diimport!`, 'success');
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload  = e => resolve(e.target.result);
    r.onerror = () => reject(new Error('Gagal membaca fail'));
    r.readAsArrayBuffer(file);
  });
}

document.getElementById('importExcelInput').addEventListener('change', function(e) {
  if (!e.target.files.length) return;
  processFiles(Array.from(e.target.files));
  e.target.value = '';
});

// Drag & drop
const importZone = document.getElementById('importZone');
importZone.addEventListener('dragover', e => { e.preventDefault(); importZone.classList.add('drag-over'); });
importZone.addEventListener('dragleave', () => importZone.classList.remove('drag-over'));
importZone.addEventListener('drop', e => {
  e.preventDefault();
  importZone.classList.remove('drag-over');
  const files = Array.from(e.dataTransfer.files).filter(f => /\.xlsx?$/i.test(f.name));
  if (files.length) processFiles(files);
});

/* =====================================================
   22. POPULATE FILTERS
   ===================================================== */
function populateFilters() {
  // Jabatan
  const jabatan = getUniqueJabatan();
  const jabOpts = jabatan.map(j=>`<option value="${j}">${j}</option>`).join('');
  ['filterJabatanPage','filterSenaraiJabatan','filterDitundaJabatan'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = `<option value="all">Semua Jabatan/Unit</option>${jabOpts}`;
  });

  // Bulan — kemaskini dropdown berdasarkan data sebenar
  const monthsWithData = getMonthsWithData();
  populateMonthDropdown('filterBulan', monthsWithData, true);   // Laporan Bulanan
  populateMonthDropdown('filterJabatanBulan', monthsWithData, false);  // Analisis Jabatan
  populateMonthDropdown('filterDitundaBulan', monthsWithData, false);  // Aktiviti Ditunda

  // Suku — kemaskini dropdown
  populateSukuDropdown();
}

function populateMonthDropdown(elId, monthsWithData, showUnavailable) {
  const el = document.getElementById(elId);
  if (!el) return;
  const allLabel = elId === 'filterBulan' ? 'Semua Bulan (Data Tersedia)' : 'Semua';
  let html = `<option value="all">${allLabel}</option>`;
  for (let m = 1; m <= 12; m++) {
    const hasData = monthsWithData.includes(m);
    if (hasData) {
      html += `<option value="${m}">${MONTH_NAMES[m-1]} 2026</option>`;
    } else if (showUnavailable) {
      html += `<option value="${m}" class="opt-unavail">${MONTH_NAMES[m-1]} 2026 — Data belum dibekalkan</option>`;
    }
  }
  el.innerHTML = html;
}

function populateSukuDropdown() {
  const el = document.getElementById('filterSuku');
  if (!el) return;
  const qWithData = getQuartersWithData();
  const qDefs = [
    { val:'Q1', label:'Q1 (Jan–Mac) 2026' },
    { val:'Q2', label:'Q2 (Apr–Jun) 2026' },
    { val:'Q3', label:'Q3 (Jul–Sep) 2026' },
    { val:'Q4', label:'Q4 (Okt–Dis) 2026' },
  ];
  let html = `<option value="all">Semua Suku (Data Tersedia)</option>`;
  qDefs.forEach(q => {
    if (qWithData.includes(q.val)) {
      html += `<option value="${q.val}">${q.label}</option>`;
    }
  });
  el.innerHTML = html;
}

/* =====================================================
   23. EKSPORT EXCEL
   ===================================================== */
function exportToExcel(rows, filename, sheetName) {
  if (!rows.length) { showToast('Tiada data untuk dieksport.','warn'); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName||'Data');
  XLSX.writeFile(wb, filename);
  showToast(`Berjaya dieksport: ${filename}`, 'success');
}

document.getElementById('btnExportAllExcel').addEventListener('click', ()=>{
  const rows = APP.data.map(a=>({
    ID: a.ID, Hari: a.Hari, Tarikh: a.TarikhDisplay,
    'Program/Aktiviti': a.NamaAktiviti, Peserta: a.Peserta,
    'Jabatan/Unit': a.Jabatan, 'Pegawai Urusetia (PIC)': a.PIC,
    'Status (Asal)': a.StatusAsal, 'Status (Normal)': a.StatusNormal,
    'Catatan/Justifikasi': a.Catatan, Sumber: a.Sumber,
    BulanPelaporan: a.TarikhISO ? MONTH_NAMES[new Date(a.TarikhISO).getMonth()] : '',
    SukuTahun: a.TarikhISO ? getQuarter(new Date(a.TarikhISO).getMonth()+1) : '',
    IsuKualiti: (a.IsuKualiti||[]).join('; '),
  }));
  exportToExcel(rows, 'Aktiviti_PBU_Q1_2026_Gabungan.xlsx', 'Aktiviti');
});

document.getElementById('btnExportBulananExcel').addEventListener('click', ()=>{
  const ms = computeMonthlyStats(APP.data);
  const rows = [1,2,3].map(m=>({
    Bulan: MONTH_NAMES[m-1]+' 2026', JumlahRekod: ms[m].total,
    Selesai: ms[m].selesai, Ditunda: ms[m].ditunda, Digugurkan: ms[m].digugurkan,
    '%Selesai': ms[m].pctSelesai, '%Ditunda': ms[m].pctDitunda, '%Digugurkan': ms[m].pctDigugurkan,
  }));
  exportToExcel(rows, 'Laporan_Bulanan_PBU_Q1_2026.xlsx', 'Bulanan');
});

document.getElementById('btnExportSukuExcel').addEventListener('click', ()=>{
  const q1 = APP.data.filter(a=>getQuarter(getMonth(a))==='Q1');
  const sq1 = computeStats(q1);
  const rows = [
    { Suku:'Q1 (Januari–Mac 2026)', JumlahRekod:sq1.total, Selesai:sq1.selesai,
      Ditunda:sq1.ditunda, Digugurkan:sq1.digugurkan,
      '%Selesai':sq1.pctSelesai, '%Ditunda':sq1.pctDitunda, '%Digugurkan':sq1.pctDigugurkan,
      StatusData:'Data tersedia' },
    { Suku:'Q2 (April–Jun 2026)',   JumlahRekod:'N/A', Selesai:'N/A',
      Ditunda:'N/A', Digugurkan:'N/A', '%Selesai':'N/A', '%Ditunda':'N/A', '%Digugurkan':'N/A', StatusData:'Data belum dibekalkan' },
    { Suku:'Q3 (Julai–Sep 2026)',   JumlahRekod:'N/A', Selesai:'N/A',
      Ditunda:'N/A', Digugurkan:'N/A', '%Selesai':'N/A', '%Ditunda':'N/A', '%Digugurkan':'N/A', StatusData:'Data belum dibekalkan' },
    { Suku:'Q4 (Okt–Dis 2026)',     JumlahRekod:'N/A', Selesai:'N/A',
      Ditunda:'N/A', Digugurkan:'N/A', '%Selesai':'N/A', '%Ditunda':'N/A', '%Digugurkan':'N/A', StatusData:'Data belum dibekalkan' },
  ];
  exportToExcel(rows, 'Laporan_Suku_Q1_PBU_2026.xlsx', 'SukuTahunan');
});

document.getElementById('btnExportJabatanExcel').addEventListener('click', ()=>{
  const rows = computeJabatanStats(APP.data).map(j=>({
    'Jabatan/Unit': j.jabatan, JumlahRekod: j.total,
    Selesai: j.selesai, Ditunda: j.ditunda, Digugurkan: j.digugurkan,
    '%Selesai': j.pctSelesai, '%Ditunda': j.pctDitunda, '%Digugurkan': j.pctDigugurkan,
    'Pegawai Urusetia (PIC)': j.pics,
  }));
  exportToExcel(rows, 'Analisis_Jabatan_PBU_Q1_2026.xlsx', 'Jabatan');
});

document.getElementById('btnExportDitundaExcel').addEventListener('click', ()=>{
  const rows = APP.data.filter(a=>a.StatusNormal==='Ditunda').map(a=>({
    ID: a.ID, 'Program/Aktiviti': a.NamaAktiviti, 'Jabatan/Unit': a.Jabatan,
    Tarikh: a.TarikhDisplay, 'Pegawai Urusetia (PIC)': a.PIC,
    'Catatan/Justifikasi': a.Catatan, 'Status Asal': a.StatusAsal,
    BulanPelaporan: a.TarikhISO ? MONTH_NAMES[new Date(a.TarikhISO).getMonth()] : '',
  }));
  exportToExcel(rows, 'Aktiviti_Ditunda_PBU_Q1_2026.xlsx', 'Ditunda');
});

document.getElementById('btnExportExcelSenarai').addEventListener('click', ()=>{
  const rows = filterSenaraiData().map(a=>({
    ID: a.ID, Hari: a.Hari, Tarikh: a.TarikhDisplay,
    'Program/Aktiviti': a.NamaAktiviti, Peserta: a.Peserta,
    'Jabatan/Unit': a.Jabatan, 'PIC': a.PIC,
    'Status': a.StatusNormal, 'Catatan': a.Catatan,
  }));
  exportToExcel(rows, 'Senarai_Aktiviti_PBU_Q1_2026.xlsx', 'Aktiviti');
});

/* =====================================================
   24. EKSPORT PDF
   ===================================================== */
function buildPDFHeader(doc, subtitle) {
  const now = new Date().toLocaleDateString('ms-MY', {day:'2-digit',month:'long',year:'numeric'});
  doc.setFontSize(14); doc.setFont(undefined,'bold'); doc.setTextColor(13,42,78);
  doc.text('POLITEKNIK BALIK PULAU', 148.5, 14, { align:'center' });
  doc.setFontSize(11); doc.setFont(undefined,'normal'); doc.setTextColor(21,101,192);
  doc.text('Laporan Pelaksanaan Aktiviti 2026 — Q1 (Januari – Mac)', 148.5, 21, { align:'center' });
  doc.setFontSize(9); doc.setFont(undefined,'normal'); doc.setTextColor(87,96,106);
  doc.text(subtitle, 148.5, 27, { align:'center' });
  doc.text(`Tarikh Jana: ${now}  |  Jumlah Rekod: ${APP.data.length}`, 148.5, 33, { align:'center' });
  doc.setDrawColor(13,42,78); doc.setLineWidth(0.5);
  doc.line(14, 36, 283, 36);
  return 40;
}

function addPDFFooter(doc) {
  const pg = doc.internal.getNumberOfPages();
  for (let i=1; i<=pg; i++) {
    doc.setPage(i);
    doc.setFontSize(8); doc.setTextColor(130);
    doc.text(`Politeknik Balik Pulau  |  Laporan Pelaksanaan Aktiviti 2026  |  Halaman ${i}/${pg}`, 148.5, 205, { align:'center' });
  }
}

function exportPDFLaporan(type) {
  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation:'landscape', unit:'mm', format:'a4' });
    let y = buildPDFHeader(doc, type === 'bulanan' ? 'Laporan Bulanan Q1 2026' :
                               type === 'ditunda'  ? 'Senarai Aktiviti Ditunda Q1 2026' :
                                                     'Laporan Pengurusan Q1 2026');

    const st = computeStats(APP.data);
    const ms = computeMonthlyStats(APP.data);

    // KPI Summary
    doc.setFontSize(9); doc.setFont(undefined,'bold'); doc.setTextColor(13,42,78);
    doc.text('RINGKASAN KPI:', 14, y); y += 6;
    doc.setFont(undefined,'normal'); doc.setTextColor(30,30,30);
    doc.text(`Jumlah Rekod: ${st.total}   |   Selesai: ${st.selesai} (${st.pctSelesai})   |   Ditunda: ${st.ditunda} (${st.pctDitunda})   |   Digugurkan: ${st.digugurkan} (${st.pctDigugurkan})   |   Jabatan/Unit: ${getUniqueJabatan().length}`, 14, y);
    y += 5;
    doc.setFontSize(8); doc.setTextColor(87,96,106);
    doc.text('Nota: Peratus dikira daripada jumlah semua rekod. Status berdasarkan rekod sumber dan tertakluk kepada pengesahan.', 14, y);
    y += 6;

    if (type === 'pengurusan' || type === 'bulanan') {
      // Jadual bulanan
      doc.setFontSize(9); doc.setFont(undefined,'bold'); doc.setTextColor(13,42,78);
      doc.text('ANALISIS BULANAN Q1 2026:', 14, y); y += 3;
      doc.autoTable({
        startY: y,
        head: [['Bulan','Jumlah Rekod','Selesai','Ditunda','Digugurkan','% Selesai','% Ditunda','% Digugurkan']],
        body: [1,2,3].map(m=>[
          MONTH_NAMES[m-1]+' 2026', ms[m].total, ms[m].selesai,
          ms[m].ditunda, ms[m].digugurkan, ms[m].pctSelesai, ms[m].pctDitunda, ms[m].pctDigugurkan
        ]),
        styles: { fontSize:9, cellPadding:3 },
        headStyles: { fillColor:[13,42,78], textColor:255, fontStyle:'bold' },
        alternateRowStyles: { fillColor:[244,246,249] },
      });
      y = doc.lastAutoTable.finalY + 8;
    }

    if (type === 'pengurusan') {
      // Analisis jabatan
      doc.setFontSize(9); doc.setFont(undefined,'bold'); doc.setTextColor(13,42,78);
      doc.text('ANALISIS JABATAN/UNIT:', 14, y); y += 3;
      const jabs = computeJabatanStats(APP.data);
      doc.autoTable({
        startY: y,
        head: [['Jabatan/Unit','Jumlah','Selesai','Ditunda','Digugurkan','% Selesai','% Ditunda']],
        body: jabs.map(j=>[j.jabatan, j.total, j.selesai, j.ditunda, j.digugurkan, j.pctSelesai, j.pctDitunda]),
        styles: { fontSize:8, cellPadding:2 },
        headStyles: { fillColor:[13,42,78], textColor:255, fontStyle:'bold' },
        alternateRowStyles: { fillColor:[244,246,249] },
        columnStyles: { 0:{ cellWidth:55 } },
      });
      y = doc.lastAutoTable.finalY + 8;

      // Senarai ditunda
      const ditunda = APP.data.filter(a=>a.StatusNormal==='Ditunda');
      if (ditunda.length) {
        doc.setFontSize(9); doc.setFont(undefined,'bold'); doc.setTextColor(13,42,78);
        doc.text(`AKTIVITI DITUNDA (${ditunda.length} rekod):`, 14, y); y += 3;
        doc.autoTable({
          startY: y,
          head: [['Program/Aktiviti','Jabatan/Unit','Tarikh','PIC','Catatan']],
          body: ditunda.map(a=>[
            String(a.NamaAktiviti||'').substring(0,40),
            String(a.Jabatan||'').substring(0,25),
            a.TarikhDisplay||'—', a.PIC||'—',
            String(a.Catatan||'').substring(0,50)
          ]),
          styles: { fontSize:7.5, cellPadding:2 },
          headStyles: { fillColor:[180,83,9], textColor:255, fontStyle:'bold' },
          alternateRowStyles: { fillColor:[254,243,199] },
          columnStyles: { 0:{ cellWidth:60 }, 1:{ cellWidth:45 }, 4:{ cellWidth:70 } },
        });
        y = doc.lastAutoTable.finalY + 8;
      }

      // Isu kualiti data
      if (APP.qualityLog.statusIssues.length) {
        doc.setFontSize(9); doc.setFont(undefined,'bold'); doc.setTextColor(13,42,78);
        doc.text('ISU KUALITI DATA:', 14, y); y += 3;
        doc.autoTable({
          startY: y,
          head: [['Program/Aktiviti','Jabatan','Status','Isu Dikenal Pasti']],
          body: APP.qualityLog.statusIssues.map(({rec,issues})=>[
            String(rec.NamaAktiviti||'').substring(0,35),
            String(rec.Jabatan||'').substring(0,25),
            rec.StatusNormal||'—',
            issues.join('; ').substring(0,70)
          ]),
          styles: { fontSize:7.5, cellPadding:2 },
          headStyles: { fillColor:[220,38,38], textColor:255, fontStyle:'bold' },
          alternateRowStyles: { fillColor:[254,226,226] },
          columnStyles: { 3:{ cellWidth:85 } },
        });
        y = doc.lastAutoTable.finalY + 8;
      }
    }

    if (type === 'ditunda') {
      const ditunda = APP.data.filter(a=>a.StatusNormal==='Ditunda');
      doc.autoTable({
        startY: y,
        head: [['Bil','Program/Aktiviti','Jabatan/Unit','Tarikh','PIC','Catatan/Justifikasi']],
        body: ditunda.map((a,i)=>[
          i+1,
          String(a.NamaAktiviti||'').substring(0,45),
          String(a.Jabatan||'').substring(0,28),
          a.TarikhDisplay||'—', a.PIC||'—',
          String(a.Catatan||'').substring(0,55)
        ]),
        styles: { fontSize:8, cellPadding:2 },
        headStyles: { fillColor:[180,83,9], textColor:255, fontStyle:'bold' },
        alternateRowStyles: { fillColor:[254,243,199] },
        columnStyles: { 1:{ cellWidth:60 }, 2:{ cellWidth:45 }, 5:{ cellWidth:75 } },
      });
    }

    // Semua aktiviti (untuk laporan bulanan)
    if (type === 'bulanan') {
      const sel = document.getElementById('filterBulan').value;
      const data = sel !== 'all'
        ? APP.data.filter(a=>getMonth(a)===parseInt(sel)) : APP.data;
      doc.autoTable({
        startY: y,
        head: [['ID','Program/Aktiviti','Jabatan/Unit','Tarikh','PIC','Status','Catatan']],
        body: data.map(a=>[
          a.ID,
          String(a.NamaAktiviti||'').substring(0,40),
          String(a.Jabatan||'').substring(0,25),
          a.TarikhDisplay||'—', a.PIC||'—',
          a.StatusNormal||'—',
          String(a.Catatan||'').substring(0,40)
        ]),
        styles: { fontSize:7.5, cellPadding:2 },
        headStyles: { fillColor:[13,42,78], textColor:255, fontStyle:'bold' },
        alternateRowStyles: { fillColor:[244,246,249] },
        columnStyles: { 1:{ cellWidth:58 }, 2:{ cellWidth:45 }, 6:{ cellWidth:55 } },
      });
    }

    addPDFFooter(doc);
    const now2 = new Date().toISOString().slice(0,10);
    const fn = { pengurusan:'Laporan_Pengurusan', bulanan:'Laporan_Bulanan', ditunda:'Laporan_Ditunda' }[type];
    doc.save(`${fn}_PBU_Q1_2026_${now2}.pdf`);
    showToast('Laporan PDF berjaya dijana.', 'success');
  } catch(err) {
    showToast('Ralat jana PDF: ' + err.message, 'error');
    console.error(err);
  }
}

document.getElementById('btnExportPDFPengurusan').addEventListener('click', ()=>exportPDFLaporan('pengurusan'));
document.getElementById('btnExportPDFBulanan').addEventListener('click',    ()=>exportPDFLaporan('bulanan'));
document.getElementById('btnExportPDFDitunda').addEventListener('click',    ()=>exportPDFLaporan('ditunda'));
document.getElementById('btnExportPDF').addEventListener('click',           ()=>exportPDFLaporan('bulanan'));

/* =====================================================
   25. PADAM DATA
   ===================================================== */
document.getElementById('btnClearData').addEventListener('click', ()=>{
  if (!confirm('Padam SEMUA data? Tindakan ini tidak boleh dibuat asal.')) return;
  [STORAGE_KEY, STORAGE_TS, STORAGE_QL, STORAGE_CNT].forEach(k=>localStorage.removeItem(k));
  APP.data = [];
  APP.hasData = false;
  APP.lastUpdated = null;
  APP.filesImported = 0;
  APP.qualityLog = { totalRaw:0, totalAfterMerge:0, emptyFields:[], statusIssues:[], duplicates:[], corrections:[], dateIssues:[] };
  document.getElementById('importLog').style.display = 'none';
  refreshAll();
  showToast('Semua data telah dipadam.', 'warn');
});

/* =====================================================
   26. NAVIGASI
   ===================================================== */
function navigateTo(page) {
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const pg  = document.getElementById('page-'+page);
  if (pg) pg.classList.add('active');
  const nav = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (nav) nav.classList.add('active');
  APP.currentPage = 1;

  switch(page) {
    case 'dashboard':  renderDashboard(); break;
    case 'bulanan':    renderBulanan();   break;
    case 'suku':       renderSuku();      break;
    case 'jabatan':    renderJabatan();   break;
    case 'ditunda':    renderDitunda();   break;
    case 'kualiti':    renderKualiti();   break;
    case 'senarai':    renderSenarai();   break;
    case 'pengurusan': updateStatusPanel(); break;
  }
  if (window.innerWidth <= 900) {
    document.getElementById('sidebar').classList.remove('open');
  }
}

document.querySelectorAll('.nav-item').forEach(item=>{
  item.addEventListener('click', e=>{ e.preventDefault(); navigateTo(item.dataset.page); });
});
document.getElementById('menuToggle').addEventListener('click', ()=>{
  document.getElementById('sidebar').classList.toggle('open');
});

/* =====================================================
   27. FILTER EVENT LISTENERS
   ===================================================== */
document.getElementById('filterBulan').addEventListener('change', renderBulanan);
document.getElementById('filterSuku').addEventListener('change', renderSuku);
['filterJabatanPage','filterJabatanBulan'].forEach(id=>{
  document.getElementById(id).addEventListener('change', renderJabatan);
});
['filterDitundaJabatan','filterDitundaBulan'].forEach(id=>{
  document.getElementById(id).addEventListener('change', renderDitunda);
});
['cariAktiviti','filterSenaraiJabatan','filterSenaraiStatus',
 'filterSenaraiSuku','filterSenaraiUrutan'].forEach(id=>{
  document.getElementById(id).addEventListener('input',  ()=>{ APP.currentPage=1; renderSenarai(); });
  document.getElementById(id).addEventListener('change', ()=>{ APP.currentPage=1; renderSenarai(); });
});

/* =====================================================
   28. PRINT
   ===================================================== */
function printPage(pageId, metaId, metaText) {
  const metaEl = document.getElementById(metaId);
  if (metaEl) {
    const now = new Date().toLocaleDateString('ms-MY', {day:'2-digit',month:'long',year:'numeric'});
    metaEl.textContent = `${metaText}  |  Tarikh Jana: ${now}`;
  }
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('print-active'));
  const target = document.getElementById(pageId);
  if (target) target.classList.add('print-active');
  setTimeout(()=>{
    window.print();
    const cleanup = ()=>{ document.querySelectorAll('.page').forEach(p=>p.classList.remove('print-active')); window.removeEventListener('afterprint',cleanup); };
    window.addEventListener('afterprint', cleanup);
    setTimeout(cleanup, 3000);
  }, 300);
}

document.getElementById('btnPrintBulanan').addEventListener('click', ()=>{
  const sel = document.getElementById('filterBulan').value;
  const label = sel==='all' ? 'Semua Bulan Q1 2026' : `Bulan: ${MONTH_NAMES[parseInt(sel)-1]} 2026`;
  printPage('page-bulanan','printMetaBulanan', label);
});
document.getElementById('btnPrintSuku').addEventListener('click', ()=>{
  printPage('page-suku','printMetaSuku','Q1 (Januari – Mac) 2026');
});

/* =====================================================
   29. REFRESH SEMUA PAPARAN
   ===================================================== */
function refreshAll() {
  populateFilters();
  updateStatusPanel();
  const activePage = document.querySelector('.page.active');
  const id = activePage ? activePage.id.replace('page-','') : 'dashboard';
  navigateTo(id);
}

/* =====================================================
   30. INIT
   ===================================================== */
function init() {
  loadFromStorage();
  populateFilters();
  updateStatusPanel();
  navigateTo('dashboard');
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
