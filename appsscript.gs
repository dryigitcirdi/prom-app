// ============================================================
// PROM Takip Sistemi – Google Apps Script Backend
// Omuz Kalsifik Tendinit | VAS · ROM · DASH
// Bu kodu Google Apps Script'e yapıştırın ve deploy edin.
// ============================================================

// Tanıya göre çizelge sayfası — her tanının verileri kendi sayfasında tutulur
const DIAG_SHEETS = {
  'Kalsifik Tendinit': 'PROM_Kalsifik_Tendinit',
  'Rotator Manşet':    'PROM_Rotator_Manset',
  'Bankart Lezyonu':   'PROM_Bankart',
};
const DEFAULT_SHEET = 'PROM_Kayitlar'; // eski kayıtlar + bilinmeyen tanılar

// Kalça artroskopisi ayrı bir modül: kendi sayfası ve kendi PROM sütunları
const HIP_SHEET = 'PROM_Kalca_Artroskopisi';
const HIP_HEADERS = [
  'Zaman Damgası', 'Tarih', 'Saat', 'Hasta ID', 'Hasta Adı', 'Tanı',
  'Girişim Tarihi', 'Post-op Gün', 'Taraf', 'Takip',
  'VAS (0-10)', 'iHOT-12 (0-100)', 'mHHS (0-100)', 'HOS-ADL (0-100)', 'HOS-Spor (0-100)',
  'Hip-RSI (0-100)', 'SF-12 PCS', 'SF-12 MCS',
];

function isHip(data) {
  return data.module === 'hip' || data.diagnosis === 'Kalça Artroskopisi';
}

// Kalça cerrahi kaydı (doktor girişi) — hasta başına tek satır
const HIP_OP_SHEET = 'Kalca_Cerrahi_Kayit';
// Sütun blokları: kimlik → PRE-OP değerlendirme → İNTRA-OP bulgular → pre-op PROM
const HIP_OP_HEADERS = [
  'Zaman Damgası', 'Hasta ID', 'Hasta Adı', 'Yaş', 'Taraf', 'Cerrahi Tarihi',
  // Pre-op (radyografi / MR)
  'Lateral CE (°)', 'CE Sınıflaması', 'Dunn Alfa (°)', 'Alfa Sınıflaması', 'FAI Tipi',
  'Kondrolabral BD — MR',
  // İntra-op (artroskopik)
  'Traksiyon (dk)', 'Kondrolabral BD — Artroskopik', 'MR / Artroskopi Uyumu',
  'Yırtık Başlangıç', 'Yırtık Bitiş', 'Yırtık Süresi (saat)',
  'Ankor Sayısı', 'Ankorlar (saat · tip)', 'Knotless', 'Düğümlü', 'Labrum İşlemi',
  'Notlar',
];

function isHipOp(data) { return data.module === 'hipop'; }

function saveHipOpRow(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HIP_OP_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HIP_OP_SHEET);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(20, 260);            // Ankorlar dökümü
    sheet.getRange('P:Q').setNumberFormat('@'); // Yırtık başlangıç/bitiş = metin
  }
  if (sheet.getRange(1, HIP_OP_HEADERS.length).getValue() !== HIP_OP_HEADERS[HIP_OP_HEADERS.length - 1]) {
    sheet.getRange(1, 1, 1, HIP_OP_HEADERS.length).setValues([HIP_OP_HEADERS]);
    const hr = sheet.getRange(1, 1, 1, HIP_OP_HEADERS.length);
    hr.setBackground('#1e6fff'); hr.setFontColor('#ffffff');
    hr.setFontWeight('bold'); hr.setFontSize(11);
  }

  const n = v => (v !== undefined && v !== null && v !== '') ? Number(v) : '';
  const row = [
    data.timestamp || '', data.patientId || '', data.patientName || '', n(data.age),
    data.side || '', data.surgeryDate || '',
    // Pre-op
    n(data.ce), data.ceClass || '', n(data.alpha), data.alphaClass || '', data.fai || '',
    data.clbPre || '',
    // İntra-op
    n(data.traction), data.clbIntra || '', data.clbConcord || '',
    data.tearStart || '', data.tearEnd || '', n(data.tearHours),
    n(data.anchorCount), data.anchors || '', n(data.knotless), n(data.knotted), data.proc || '',
    data.notes || '',
  ];

  // Aynı hasta zaten varsa satırı güncelle, yoksa ekle
  const last = sheet.getLastRow();
  let target = 0;
  if (last >= 2) {
    const vals = sheet.getRange(2, 2, last - 1, 2).getValues(); // Hasta ID + Ad
    const id = String(data.patientId || '').trim();
    const nm = String(data.patientName || '').trim().toLowerCase();
    for (let i = 0; i < vals.length; i++) {
      const rid = String(vals[i][0] || '').trim();
      const rnm = String(vals[i][1] || '').trim().toLowerCase();
      if ((id && rid) ? rid === id : (nm && rnm === nm)) { target = i + 2; break; }
    }
  }
  if (target) sheet.getRange(target, 1, 1, row.length).setValues([row]);
  else { sheet.appendRow(row); target = sheet.getLastRow(); }

  // Yırtık saatleri ("12:30") Sheets tarafından saat değerine çevrilmesin — düz metin
  sheet.getRange(target, 16, 1, 2)
    .setNumberFormat('@')
    .setValues([[String(data.tearStart || ''), String(data.tearEnd || '')]]);

  // Sütunlar (1-based): CE=7, Alfa=9, KLB-MR=12, KLB-Artro=14, Uyum=15
  colorCE(sheet, target, 7, data.ce);
  colorAlpha(sheet, target, 9, data.alpha);
  colorCLB(sheet, target, 14, data.clbIntra);
  colorConcord(sheet, target, 15, data.clbConcord);
}

// Lateral CE: <20 displazi (kırmızı) · 20-25 borderline (sarı) · 25-40 normal (yeşil) · >40 overcoverage (turuncu)
function colorCE(sheet, row, col, v) {
  if (v === undefined || v === null || v === '') return;
  v = Number(v);
  const bg = v < 20 ? '#fee2e2' : v <= 25 ? '#fef9c3' : v <= 40 ? '#dcfce7' : '#ffedd5';
  sheet.getRange(row, col).setBackground(bg);
}

// Kondrolabral breakdown (artroskopik): Var = sarı, Yok = yeşil
function colorCLB(sheet, row, col, v) {
  if (!v) return;
  sheet.getRange(row, col).setBackground(v === 'Var' ? '#fef9c3' : '#dcfce7');
}

// MR / artroskopi uyumu: MR'da kaçmışsa kırmızı, uyumluysa yeşil/sarı
function colorConcord(sheet, row, col, v) {
  if (!v) return;
  const bg = v.indexOf('görülmedi') > -1 ? '#fee2e2'
           : v.indexOf('intra-op yok') > -1 ? '#ffedd5'
           : v.indexOf('uyumlu (yok)') > -1 ? '#dcfce7' : '#fef9c3';
  sheet.getRange(row, col).setBackground(bg);
}

// Dunn alfa: ≤55 normal · 55-60 sınırda · >60 CAM
function colorAlpha(sheet, row, col, v) {
  if (v === undefined || v === null || v === '') return;
  v = Number(v);
  const bg = v <= 55 ? '#dcfce7' : v <= 60 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

// ═══ Tümör modülü — 3 sayfa: hasta künyesi · olaylar · lezyon çizimleri ═══
const TUMOR_SHEET  = 'Tumor_Hastalar';
const TUMOR_EVENTS = 'Tumor_Olaylar';
const TUMOR_DRAWS  = 'Tumor_Cizimler';

const TUMOR_HEADERS = [
  'Zaman Damgası', 'Hasta ID', 'Hasta Adı', 'Yaş', 'Tümör Tipi',
  'Ön Tanı', 'Biyopsi', 'Kesin Tanı',
  'Lokalizasyon', 'Bölge Kodları',
  'Cerrahi Tarihi', 'Cerrahi Sınır',
  'Nüks', 'Nüks Tarihi', 'Nüks Süresi (ay)',
  'Metastaz', 'Metastaz Bölgeleri',
  'Olay Sayısı', 'Çizim Sayısı', 'Notlar',
];
const TUMOR_EVENT_HEADERS = [
  'Zaman Damgası', 'Hasta ID', 'Hasta Adı', 'Olay Tarihi', 'Olay Tipi', 'Bölgeler', 'Not',
];
const TUMOR_DRAW_HEADERS = [
  'Zaman Damgası', 'Hasta ID', 'Hasta Adı', 'Bölge', 'Çizim (PNG)',
];

function isTumor(data) { return data.module === 'tumor'; }

function ensureSheet(name, headers, widths) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.setFrozenRows(1);
    (widths || []).forEach(function (w) { sheet.setColumnWidth(w[0], w[1]); });
  }
  if (sheet.getRange(1, headers.length).getValue() !== headers[headers.length - 1]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground('#1e6fff'); hr.setFontColor('#ffffff');
    hr.setFontWeight('bold'); hr.setFontSize(11);
  }
  return sheet;
}

// Bir hastanın mevcut satırlarını sil (olaylar/çizimler yeniden yazılacak)
function clearPatientRows(sheet, idCol, id, nameCol, name) {
  const last = sheet.getLastRow();
  if (last < 2) return;
  const vals = sheet.getRange(2, 1, last - 1, Math.max(idCol, nameCol)).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const rid = String(vals[i][idCol - 1] || '').trim();
    const rnm = String(vals[i][nameCol - 1] || '').trim().toLowerCase();
    const match = (id && rid) ? rid === id : (name && rnm === name);
    if (match) sheet.deleteRow(i + 2);
  }
}

function saveTumorData(data) {
  const id = String(data.patientId || '').trim();
  const nameLc = String(data.patientName || '').trim().toLowerCase();
  const n = v => (v !== undefined && v !== null && v !== '') ? Number(v) : '';

  // ── 1) Hasta künyesi (hasta başına tek satır: varsa güncelle) ──
  const sheet = ensureSheet(TUMOR_SHEET, TUMOR_HEADERS, [[1, 160], [3, 160], [9, 260], [17, 200]]);
  const row = [
    data.timestamp || '', data.patientId || '', data.patientName || '', n(data.age),
    data.ttype || '', data.preDx || '', data.biopsy || '', data.finalDx || '',
    data.regionLabels || '', data.regionIds || '',
    data.surgery || '', data.margin || '',
    data.rec || '', data.recDate || '', n(data.recMonths),
    (data.mets && data.mets.length) ? 'Var' : 'Yok', data.metLabels || '',
    n(data.eventCount), n(data.drawCount), data.notes || '',
  ];
  const last = sheet.getLastRow();
  let target = 0;
  if (last >= 2) {
    const vals = sheet.getRange(2, 2, last - 1, 2).getValues();
    for (let i = 0; i < vals.length; i++) {
      const rid = String(vals[i][0] || '').trim();
      const rnm = String(vals[i][1] || '').trim().toLowerCase();
      if ((id && rid) ? rid === id : (nameLc && rnm === nameLc)) { target = i + 2; break; }
    }
  }
  if (target) sheet.getRange(target, 1, 1, row.length).setValues([row]);
  else { sheet.appendRow(row); target = sheet.getLastRow(); }
  colorMargin(sheet, target, 12, data.margin);
  colorRec(sheet, target, 13, data.rec);
  colorMet(sheet, target, 16, (data.mets && data.mets.length) ? 'Var' : 'Yok');

  // ── 2) Zaman çizelgesi olayları (hastanın satırları yeniden yazılır) ──
  if (data.events) {
    const ev = ensureSheet(TUMOR_EVENTS, TUMOR_EVENT_HEADERS, [[1, 160], [3, 160], [6, 240], [7, 320]]);
    clearPatientRows(ev, 2, id, 3, nameLc);
    data.events.forEach(function (e) {
      ev.appendRow([data.timestamp || '', data.patientId || '', data.patientName || '',
                    e.date || '', e.typeLabel || e.type || '', e.regionLabels || '', e.note || '']);
    });
  }

  // ── 3) Lezyon çizimleri ──
  if (data.drawings) {
    const dr = ensureSheet(TUMOR_DRAWS, TUMOR_DRAW_HEADERS, [[1, 160], [3, 160], [4, 200], [5, 420]]);
    clearPatientRows(dr, 2, id, 3, nameLc);
    data.drawings.forEach(function (d) {
      const png = String(d.png || '');
      dr.appendRow([data.timestamp || '', data.patientId || '', data.patientName || '',
                    d.label || '', png.length > 45000 ? '[çok büyük — yerel]' : png]);
    });
  }
}

// Cerrahi sınır: Temiz yeşil · Marjinal sarı · Pozitif kırmızı
function colorMargin(sheet, row, col, v) {
  if (!v) return;
  const bg = v === 'Temiz' ? '#dcfce7' : v === 'Marjinal' ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}
function colorRec(sheet, row, col, v) {
  if (!v) return;
  sheet.getRange(row, col).setBackground(v === 'Var' ? '#fee2e2' : '#dcfce7');
}
function colorMet(sheet, row, col, v) {
  if (!v) return;
  sheet.getRange(row, col).setBackground(v === 'Var' ? '#f3e8ff' : '#dcfce7');
}

function sheetNameFor(diag) {
  return DIAG_SHEETS[diag] || DEFAULT_SHEET;
}

// Takip dönemi sırası — aynı hastanın satırları tabloda bu sıraya göre alt alta dizilir
// Yaklaşık gün değeriyle sıralanır — omuz (2./6. hafta) ve kalça (1./3./12. ay) birlikte çalışır
const FU_ORDER = {
  'Pre-op': 0, '2. Hafta': 14, '1. Ay': 30, '6. Hafta': 42,
  '3. Ay': 90, '6. Ay': 180, '12. Ay': 365,
};
function fuRank(label) {
  return FU_ORDER.hasOwnProperty(label) ? FU_ORDER[label] : 9999;
}

// Yeni kaydın gideceği satır numarası: aynı hastanın bloğu içinde dönem sırasına
// göre konum bulur (hasta anahtarı: Hasta ID, yoksa ad). Hasta yoksa null → sona eklenir.
function findInsertRow(sheet, data) {
  return findInsertRowIn(sheet, data, HEADERS.length, 20);
}

function findInsertRowIn(sheet, data, colCount, fuCol) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const values = sheet.getRange(2, 1, last - 1, colCount).getValues();
  const newId   = String(data.patientId || '').trim();
  const newName = String(data.patientName || '').trim().toLowerCase();
  const newRank = fuRank(String(data.followup || ''));

  let lastMatch = -1; // hastanın son satırı (values dizini)
  let insertAt  = -1; // dönem sırası daha büyük olan ilk satırı (onun üstüne girilir)
  for (let i = 0; i < values.length; i++) {
    const id   = String(values[i][3] || '').trim();
    const name = String(values[i][4] || '').trim().toLowerCase();
    const same = (newId && id) ? id === newId : (newName !== '' && name === newName);
    if (!same) continue;
    lastMatch = i;
    if (insertAt === -1 && fuRank(String(values[i][fuCol] || '')) > newRank) insertAt = i;
  }
  if (lastMatch === -1) return null;      // yeni hasta
  if (insertAt !== -1) return insertAt + 2; // bu satırın üstüne (sheet satır no = dizin + 2)
  return lastMatch + 3;                     // bloğun hemen altına
}

// Sütun sırası (0-based index):
// 0:timestamp  1:date      2:time       3:patientId  4:patientName  5:diagnosis
// 6:surgeryDate 7:daysPostOp 8:surgeryLeg
// 9:vas  10:rom_abduction  11:rom_flexion  12:rom_ir  13:rom_ir_label  14:rom_er
// 15:dash  16:constant  17:sleep_vas  18:sf12_pcs  19:sf12_mcs  20:followup
const HEADERS = [
  'Zaman Damgası', 'Tarih', 'Saat', 'Hasta ID', 'Hasta Adı', 'Tanı',
  'Girişim Tarihi', 'Post-op Gün', 'Taraf',
  'VAS (0-10)', 'Abduksiyon (°)', 'Öne Fleksiyon (°)', 'İç Rotasyon', 'İR Pozisyon', 'Dış Rotasyon (°)',
  'DASH (0-100)', 'Constant (0-100)', 'Uyku VAS (0-10)', 'SF-12 PCS', 'SF-12 MCS', 'Takip',
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    saveToSheet(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e.parameter.data) {
    try {
      const data = JSON.parse(e.parameter.data);
      saveToSheet(data);
      return HtmlService.createHtmlOutput('<script>window.close();</script>');
    } catch (err) {
      return HtmlService.createHtmlOutput('<script>window.close();</script>');
    }
  }
  if (e.parameter.action === 'getData') {
    return getSheetData();
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'PROM API çalışıyor' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveToSheet(data) {
  if (isTumor(data)) { saveTumorData(data); return; }
  if (isHipOp(data)) { saveHipOpRow(data); return; }
  if (isHip(data)) { saveHipRow(data); return; }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sheetNameFor(data.diagnosis);
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);  // Zaman Damgası
    sheet.setColumnWidth(5, 160);  // Hasta Adı
  }

  // Başlık satırını yaz / yeni sütun eklendiğinde mevcut sayfada tamamla
  if (sheet.getRange(1, HEADERS.length).getValue() !== HEADERS[HEADERS.length - 1]) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    const hr = sheet.getRange(1, 1, 1, HEADERS.length);
    hr.setBackground('#1e6fff');
    hr.setFontColor('#ffffff');
    hr.setFontWeight('bold');
    hr.setFontSize(11);
  }

  const n = v => (v !== undefined && v !== null && v !== '') ? Number(v) : '';

  const row = [
    data.timestamp    || '',
    data.date         || '',
    data.time         || '',
    data.patientId    || '',
    data.patientName  || '',
    data.diagnosis    || '',
    data.surgeryDate  || '',
    n(data.daysPostOp),
    data.surgeryLeg   || '',
    n(data.vas),
    n(data.rom_abduction),
    n(data.rom_flexion),
    n(data.rom_ir),
    data.rom_ir_label || '',
    n(data.rom_er),
    n(data.dash),
    n(data.constant),
    n(data.sleep_vas),
    n(data.sf12_pcs),
    n(data.sf12_mcs),
    data.followup || '',
  ];

  // Aynı hastanın satırları dönem sırasıyla alt alta gelecek şekilde yerleştir
  const target = findInsertRow(sheet, data);
  let rowIdx;
  if (target === null || target > sheet.getLastRow()) {
    sheet.appendRow(row);
    rowIdx = sheet.getLastRow();
  } else {
    sheet.insertRowsBefore(target, 1);
    sheet.getRange(target, 1, 1, row.length).setValues([row]);
    rowIdx = target;
  }

  // Sütun numaraları (1-based): VAS=10, Abd=11, Flex=12, ER=15, DASH=16, Constant=17, Uyku=18, PCS=19, MCS=20
  colorVAS(sheet, rowIdx, 10, data.vas);
  colorROM(sheet, rowIdx, 11, data.rom_abduction, 150, 90);
  colorROM(sheet, rowIdx, 12, data.rom_flexion,   150, 90);
  colorROM(sheet, rowIdx, 15, data.rom_er,          60, 30);
  colorDASH(sheet, rowIdx, 16, data.dash);
  colorConstant(sheet, rowIdx, 17, data.constant);
  colorSleep(sheet, rowIdx, 18, data.sleep_vas);
  colorSF12(sheet, rowIdx, 19, data.sf12_pcs);
  colorSF12(sheet, rowIdx, 20, data.sf12_mcs);
}

// VAS: düşük = iyi
function colorVAS(sheet, row, col, score) {
  if (score === undefined || score === null || score === '') return;
  const v = Number(score);
  const bg = v <= 3 ? '#dcfce7' : v <= 6 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

// ROM: yüksek = iyi
function colorROM(sheet, row, col, score, good, mid) {
  if (score === undefined || score === null || score === '') return;
  const v = Number(score);
  const bg = v >= good ? '#dcfce7' : v >= mid ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

// DASH: düşük = iyi (0=bağımsız, 100=tam kısıtlı)
function colorDASH(sheet, row, col, score) {
  if (score === undefined || score === null || score === '') return;
  const v = Number(score);
  const bg = v <= 25 ? '#dcfce7' : v <= 50 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

// Constant: yüksek = iyi (≥80=mükemmel, ≥60=orta, <60=kötü)
function colorConstant(sheet, row, col, score) {
  if (score === undefined || score === null || score === '') return;
  const v = Number(score);
  const bg = v >= 80 ? '#dcfce7' : v >= 60 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

// Uyku VAS: yüksek = iyi (≥7=iyi, ≥4=orta, <4=kötü)
function colorSleep(sheet, row, col, score) {
  if (score === undefined || score === null || score === '') return;
  const v = Number(score);
  const bg = v >= 7 ? '#dcfce7' : v >= 4 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

// SF-12 PCS/MCS: norm 50 ± 10 (≥50=iyi, ≥40=orta, <40=kötü)
function colorSF12(sheet, row, col, score) {
  if (score === undefined || score === null || score === '') return;
  const v = Number(score);
  const bg = v >= 50 ? '#dcfce7' : v >= 40 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

// Kalça artroskopisi kaydı — kendi sayfasına, hasta bazında dönem sırasıyla
function saveHipRow(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HIP_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HIP_SHEET);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(5, 160);
  }
  if (sheet.getRange(1, HIP_HEADERS.length).getValue() !== HIP_HEADERS[HIP_HEADERS.length - 1]) {
    sheet.getRange(1, 1, 1, HIP_HEADERS.length).setValues([HIP_HEADERS]);
    const hr = sheet.getRange(1, 1, 1, HIP_HEADERS.length);
    hr.setBackground('#1e6fff'); hr.setFontColor('#ffffff');
    hr.setFontWeight('bold'); hr.setFontSize(11);
  }

  const n = v => (v !== undefined && v !== null && v !== '') ? Number(v) : '';
  const row = [
    data.timestamp || '', data.date || '', data.time || '',
    data.patientId || '', data.patientName || '', 'Kalça Artroskopisi',
    data.surgeryDate || '', n(data.daysPostOp), data.surgeryLeg || '', data.followup || '',
    n(data.vas), n(data.ihot12), n(data.mhhs), n(data.hos_adl), n(data.hos_sport),
    n(data.hip_rsi), n(data.sf12_pcs), n(data.sf12_mcs),
  ];

  // Aynı hastanın satırları dönem sırasıyla alt alta
  const target = findInsertRowIn(sheet, data, HIP_HEADERS.length, 9);
  let rowIdx;
  if (target === null || target > sheet.getLastRow()) {
    sheet.appendRow(row); rowIdx = sheet.getLastRow();
  } else {
    sheet.insertRowsBefore(target, 1);
    sheet.getRange(target, 1, 1, row.length).setValues([row]);
    rowIdx = target;
  }

  colorVAS(sheet, rowIdx, 11, data.vas);
  colorHip(sheet, rowIdx, 12, data.ihot12);
  colorHip(sheet, rowIdx, 13, data.mhhs);
  colorHip(sheet, rowIdx, 14, data.hos_adl);
  colorHip(sheet, rowIdx, 15, data.hos_sport);
  colorRSI(sheet, rowIdx, 16, data.hip_rsi);
  colorSF12(sheet, rowIdx, 17, data.sf12_pcs);
  colorSF12(sheet, rowIdx, 18, data.sf12_mcs);
}

// Hip-RSI: ≥70 hazır · ≥50 kısmen · <50 hazır değil
function colorRSI(sheet, row, col, score) {
  if (score === undefined || score === null || score === '') return;
  const v = Number(score);
  const bg = v >= 70 ? '#dcfce7' : v >= 50 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

// Kalça PROM'ları: yüksek = iyi (≥80 yeşil, ≥60 sarı, <60 kırmızı)
function colorHip(sheet, row, col, score) {
  if (score === undefined || score === null || score === '') return;
  const v = Number(score);
  const bg = v >= 80 ? '#dcfce7' : v >= 60 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(bg);
}

function getSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Tüm tanı sayfalarını tek listede birleştir (ilk satır ortak başlık)
  const names = [DEFAULT_SHEET].concat(Object.keys(DIAG_SHEETS).map(k => DIAG_SHEETS[k]));
  const rows = [HEADERS];
  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) rows.push(data[i]); // başlığı atla
  });
  // Kalça artroskopisi kayıtları ayrı sütun düzeninde
  const hipRows = [HIP_HEADERS];
  const hipSheet = ss.getSheetByName(HIP_SHEET);
  if (hipSheet && hipSheet.getLastRow() >= 2) {
    const hd = hipSheet.getDataRange().getValues();
    for (let i = 1; i < hd.length; i++) hipRows.push(hd[i]);
  }

  // Kalça cerrahi kayıtları
  const hipOpRows = [HIP_OP_HEADERS];
  const opSheet = ss.getSheetByName(HIP_OP_SHEET);
  if (opSheet && opSheet.getLastRow() >= 2) {
    const od = opSheet.getDataRange().getValues();
    for (let i = 1; i < od.length; i++) hipOpRows.push(od[i]);
  }

  // Tümör: hasta künyeleri ve olaylar (çizimler yanıta dahil edilmez — boyut)
  function dumpSheet(name, headers) {
    const out = [headers];
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() >= 2) {
      const d = sh.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) out.push(d[i]);
    }
    return out;
  }
  const tumorRows = dumpSheet(TUMOR_SHEET, TUMOR_HEADERS);
  const tumorEventRows = dumpSheet(TUMOR_EVENTS, TUMOR_EVENT_HEADERS);

  return ContentService
    .createTextOutput(JSON.stringify({ rows: rows, hipRows: hipRows, hipOpRows: hipOpRows,
                                       tumorRows: tumorRows, tumorEventRows: tumorEventRows }))
    .setMimeType(ContentService.MimeType.JSON);
}
