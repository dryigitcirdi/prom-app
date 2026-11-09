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
  'Metastaz', 'Metastaz Bölgeleri', 'Metastaz Kodları',
  'Olay Sayısı', 'Çizim Sayısı', 'Notlar',
  // Yumuşak doku tümörlerine özel (sona eklendi — mevcut sütun sırası bozulmasın)
  'Derinlik', 'Büyüklük (cm)', 'Grade', 'RT', 'RT Zamanlama', 'KT', 'KT Zamanlama',
  'Tanı Tarihi', 'Nüks Bölgeleri', 'Nüks Kodları', 'Alt Tip',
];
const TUMOR_EVENT_HEADERS = [
  'Zaman Damgası', 'Hasta ID', 'Hasta Adı', 'Olay Tarihi', 'Olay Tipi', 'Bölgeler', 'Not',
  'Olay Kodu', 'Bölge Kodları', 'Bitiş Tarihi', 'Sonuç',
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
    (data.mets || []).join(' | '),
    n(data.eventCount), n(data.drawCount), data.notes || '',
    data.depth || '', n(data.size), data.grade || '',
    data.rt || '', data.rtWhen || '', data.kt || '', data.ktWhen || '',
    data.dxDate || '', data.recLabels || '', (data.recRegions || []).join(' | '),
    data.subtype || '',
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
  colorDepth(sheet, target, 22, data.depth);
  colorGrade(sheet, target, 24, data.grade);

  // ── 2) Zaman çizelgesi olayları (hastanın satırları yeniden yazılır) ──
  if (data.events) {
    const ev = ensureSheet(TUMOR_EVENTS, TUMOR_EVENT_HEADERS, [[1, 160], [3, 160], [6, 240], [7, 320]]);
    clearPatientRows(ev, 2, id, 3, nameLc);
    data.events.forEach(function (e) {
      ev.appendRow([data.timestamp || '', data.patientId || '', data.patientName || '',
                    e.date || '', e.typeLabel || e.type || '', e.regionLabels || '', e.note || '',
                    e.type || '', (e.regionIds || []).join(' | '), e.end || '', e.outcome || '']);
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
    saveGuarded_(data);
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
  // Hasta formları yalnızca ?data= ile gelir. Hekim uygulaması ?action=aclDoc&data=…
  // kullanır ve JSON yanıt bekler — bu yüzden action varsa aşağıdaki dala düşsün.
  if (e.parameter.data && !e.parameter.action) {
    try {
      const data = JSON.parse(e.parameter.data);
      saveGuarded_(data);
      return HtmlService.createHtmlOutput('<script>window.close();</script>');
    } catch (err) {
      return HtmlService.createHtmlOutput('<script>window.close();</script>');
    }
  }
  if (e.parameter.action === 'getData') {
    if (!pinOk_(e)) return unauthorized_();
    return getSheetData();
  }
  if (e.parameter.action === 'tumorDrawings') {
    if (!pinOk_(e)) return unauthorized_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(TUMOR_DRAWS);
    const rows = [TUMOR_DRAW_HEADERS];
    if (sh && sh.getLastRow() >= 2) {
      const d = sh.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) rows.push(d[i]);
    }
    return ContentService.createTextOutput(JSON.stringify({ rows: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === 'aclDoc') {
    if (!pinOk_(e)) return unauthorized_();
    const lock = LockService.getScriptLock();
    let alindi = false;
    try { alindi = lock.tryLock(25000); } catch (err) { alindi = false; }
    try {
      const vaka = saveAclDoctorRow(JSON.parse(e.parameter.data || '{}'));
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', vaka: vaka }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    } finally {
      if (alindi) { try { lock.releaseLock(); } catch (e2) {} }
    }
  }
  if (e.parameter.action === 'aclEdit') {
    if (!pinOk_(e)) return unauthorized_();
    const lock = LockService.getScriptLock();
    let alindi = false;
    try { alindi = lock.tryLock(25000); } catch (err) { alindi = false; }
    try {
      const degisen = saveAclPatientEdit(JSON.parse(e.parameter.data || '{}'));
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', changed: degisen }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    } finally {
      if (alindi) { try { lock.releaseLock(); } catch (e2) {} }
    }
  }
  if (e.parameter.action === 'aclCases') {
    if (!pinOk_(e)) return unauthorized_();
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', cases: aclCaseList() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === 'aclCase') {
    if (!pinOk_(e)) return unauthorized_();
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', vaka: aclCaseDetail(e.parameter.vaka) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === 'newlink') {
    if (!pinOk_(e)) return unauthorized_();
    const code = makeShortLink(e.parameter);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', code: code }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === 'link') {
    const rec = resolveShortLink(e.parameter.k);
    return ContentService.createTextOutput(JSON.stringify(rec))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'PROM API çalışıyor' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Aynı hastanın iki kaydı aynı anda işlenirse satırlar çakışıyor: biri silerken
// diğeri okuyor, sonuç çift satır. Yazma işlemini betik kilidiyle sıraya sok.
function saveGuarded_(data) {
  const lock = LockService.getScriptLock();
  let alindi = false;
  try { alindi = lock.tryLock(25000); } catch (err) { alindi = false; }
  try {
    saveToSheet(data);
  } finally {
    if (alindi) { try { lock.releaseLock(); } catch (err2) {} }
  }
}

function saveToSheet(data) {
  if (data && data.module === 'rc')   { saveRcRow(data);   return; }
  if (data && data.module === 'acl-hekim') { saveAclDoctorRow(data); return; }
  if (data && data.module === 'acl')  { saveAclRow(data);  return; }
  if (data && data.module === 'uyku') { saveUykuRow(data); return; }
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
  const uykuRows = dumpSheet(UYKU_SHEET, UYKU_HEADERS);
  const aclRows = dumpSheet(ACL_SHEET, ACL_HEADERS);
  const rcRows  = dumpSheet(RC_SHEET,  RC_HEADERS);

  return ContentService
    .createTextOutput(JSON.stringify({ rows: rows, hipRows: hipRows, hipOpRows: hipOpRows,
                                       tumorRows: tumorRows, tumorEventRows: tumorEventRows,
                                       uykuRows: uykuRows, aclRows: aclRows, rcRows: rcRows }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════════════════
// KISA LİNK — hasta adı/tarihi URL'de görünmesin diye kod tabanlı yönlendirme
// ═══════════════════════════════════════════════════════════════════════════
const LINK_SHEET = 'Kisa_Linkler';
const LINK_HEADERS = ['Kod', 'Modül', 'Ad Soyad', 'Hasta ID', 'Girişim Tarihi',
                      'Taraf', 'Dönem', 'Tanı', 'Oluşturma', 'İlk Açılma', 'Açılma Sayısı'];
// Karışabilen karakterler (I, O, 0, 1) alfabede yok — telefonda okunaklı olsun
const LINK_ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

function linkSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LINK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LINK_SHEET);
    sh.getRange(1, 1, 1, LINK_HEADERS.length).setValues([LINK_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, LINK_HEADERS.length).setBackground('#1e293b').setFontColor('#ffffff');
    sh.setColumnWidth(1, 80);
  }
  return sh;
}

function makeShortLink(p) {
  const sh = linkSheet_();
  const last = sh.getLastRow();
  const used = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  let code = '';
  for (let tries = 0; tries < 40; tries++) {
    code = '';
    for (let i = 0; i < 5; i++) code += LINK_ALPHABET.charAt(Math.floor(Math.random() * LINK_ALPHABET.length));
    if (used.indexOf(code) === -1) break;
  }
  sh.appendRow([code, p.module || 'omuz', p.name || '', p.pid || '', p.surgery || '',
                p.leg || '', p.fu || '', p.procedure || '', new Date(), '', 0]);
  // Kod ve Hasta ID metin kalsın (Sheets sayıya çevirmesin)
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('@').setValue(code);
  sh.getRange(sh.getLastRow(), 4).setNumberFormat('@').setValue(String(p.pid || ''));
  return code;
}

function resolveShortLink(code) {
  const key = String(code || '').toUpperCase().trim();
  if (!key) return { status: 'error', message: 'kod yok' };
  const sh = linkSheet_();
  if (sh.getLastRow() < 2) return { status: 'error', message: 'bulunamadı' };
  const d = sh.getRange(2, 1, sh.getLastRow() - 1, LINK_HEADERS.length).getValues();
  for (let i = 0; i < d.length; i++) {
    if (String(d[i][0]).toUpperCase().trim() !== key) continue;
    const row = i + 2;
    if (!d[i][9]) sh.getRange(row, 10).setValue(new Date());
    sh.getRange(row, 11).setValue((Number(d[i][10]) || 0) + 1);
    return {
      status: 'ok',
      module: String(d[i][1] || 'omuz'),
      name: String(d[i][2] || ''),
      id: String(d[i][3] || ''),
      surgery: fmtDate_(d[i][4]),
      leg: String(d[i][5] || ''),
      fu: String(d[i][6] || ''),
      procedure: String(d[i][7] || '')
    };
  }
  return { status: 'error', message: 'bulunamadı' };
}

// Tarih hücresi Date olarak dönerse YYYY-MM-DD'ye çevir (saat dilimi kaymasın)
function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Istanbul', 'yyyy-MM-dd');
  return String(v || '');
}



// ═══════════════════════════════════════════════════════════════════════════
// DASHBOARD PIN — hasta verisini okuyan uçlar PIN ister.
// PIN kod içinde DEĞİL, Script Properties'te (DASH_PIN) durur; bu dosya
// herkese açık depoda olduğu için buraya asla yazılmamalı.
// Hasta anketi gönderimleri (doPost / ?data=) PIN istemez — hastada PIN yok.
// ═══════════════════════════════════════════════════════════════════════════
function pinOk_(e) {
  const want = PropertiesService.getScriptProperties().getProperty('DASH_PIN');
  if (!want) return true; // PIN kurulmamışsa sistemi kilitleme
  return String((e && e.parameter && e.parameter.pin) || '') === String(want);
}

function unauthorized_() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
}




// Derinlik ve grade renk kodları — derin ve yüksek grade daha riskli
function colorDepth(sheet, row, col, v) {
  const c = String(v || '') === 'Derin' ? '#fee2e2' : (String(v || '') === 'Yüzeysel' ? '#dcfce7' : null);
  sheet.getRange(row, col).setBackground(c);
}

function colorGrade(sheet, row, col, v) {
  const g = String(v || '').charAt(0);
  const map = { '1': '#dcfce7', '2': '#fef9c3', '3': '#fee2e2' };
  sheet.getRange(row, col).setBackground(map[g] || null);
}



// ═══════════════════════════════════════════════════════════════════════════
// PROM ÇALIŞMALARI — Uyku Pozisyonu ve Omuz Ağrısı
// Bu çalışma diğer modüllerden bağımsız; kendi sayfasında toplanır.
// ═══════════════════════════════════════════════════════════════════════════
const UYKU_SHEET = 'PROM_Uyku_Pozisyonu';
const UYKU_HEADERS = [
  'Zaman Damgası', 'Tarih', 'Hasta ID', 'Ad Soyad', 'Tanı',
  'Cinsiyet', 'Yaş', 'Boy (cm)', 'Kilo (kg)', 'VKİ',
  'Meslek', 'Mesai', 'Dominant El', 'Ek Hastalıklar',
  'Uyku Pozisyonu', 'Pozisyon Kodu', 'Yattığı Kol',
  'Partner', 'Yatak Tarafı',
  'Ağrılı Omuz', 'Ağrı Süresi', 'Gece Uyandırma', 'Sabah Tutukluğu',
  'Geçmiş Omuz Ağrısı', 'Başlangıç', 'Omuza Yük',
  'İzinde Ağrı', 'Hafta Sonu Ağrı', 'Ağrı Saatleri',
  'Pozisyon No', 'Kaynak',
];

function saveUykuRow(data) {
  // n() bu projede global değil, her fonksiyonda ayrı tanımlı
  const n = v => (v !== undefined && v !== null && v !== '') ? Number(v) : '';
  const sheet = ensureSheet(UYKU_SHEET, UYKU_HEADERS, [[4, 150], [5, 160], [11, 140], [14, 200], [15, 130]]);
  const ts = data.timestamp ? new Date(data.timestamp) : new Date();
  sheet.appendRow([
    ts, Utilities.formatDate(ts, 'Europe/Istanbul', 'yyyy-MM-dd'),
    data.patientId || '', data.name || '', data.dx || '',
    data.gender || '', n(data.age), n(data.height), n(data.weight), n(data.bmi),
    data.job || '', data.shift || '', data.dominant || '', data.comorbid || '',
    data.sleepPos || '', data.sleepPosCode || '', data.lieOnArm || '',
    data.partner || '', data.bedSide || '',
    data.painSide || '', data.duration || '', data.nightPain || '', data.stiffness || '',
    data.priorPain || '', data.onset || '', data.load || '',
    data.leavePain || '', data.weekendPain || '', data.painHours || '',
    data.sleepPosId || '', data.src || '',
  ]);
  const row = sheet.getLastRow();
  colorEvetHayir(sheet, row, 22, data.nightPain);
  colorEvetHayir(sheet, row, 23, data.stiffness);
  colorBmi(sheet, row, 10, data.bmi);
}

// Evet = dikkat (kırmızımsı), Hayır = sakin (yeşilimsi)
function colorEvetHayir(sheet, row, col, v) {
  const c = String(v || '') === 'Evet' ? '#fee2e2' : (String(v || '') === 'Hayır' ? '#dcfce7' : null);
  sheet.getRange(row, col).setBackground(c);
}

function colorBmi(sheet, row, col, v) {
  const b = Number(v);
  if (!b) { sheet.getRange(row, col).setBackground(null); return; }
  const c = b < 18.5 ? '#fef9c3' : b < 25 ? '#dcfce7' : b < 30 ? '#fef9c3' : '#fee2e2';
  sheet.getRange(row, col).setBackground(c);
}



// ═══════════════════════════════════════════════════════════════════════════
// ACL AMELİYAT ÖNCESİ — kayıt sistemi uyumlu veri seti
// Sütun düzeni Norveç (NKLR), Kaiser Permanente ve İsveç kayıtlarıyla
// karşılaştırılabilir olacak şekilde seçildi.
// ═══════════════════════════════════════════════════════════════════════════
const ACL_SHEET = 'PROM_ACL_PreOp';
const ACL_HEADERS = [
  'Zaman Damgası', 'Tarih', 'Hasta ID', 'Ad Soyad', 'Kaynak',
  'Yaş', 'Cinsiyet', 'Boy (cm)', 'Kilo (kg)', 'VKİ', 'Dominant Bacak', 'Sigara',
  'Yaralanma Tarihi', 'Geçen Gün', 'Taraf',
  'Spor', 'Düzey', 'Zemin', 'Ortam',
  'Haftada Kez', 'Seans (dk)', 'Yılda Ay', 'Kaç Yıldır', 'Haftalık Saat', 'Toplam Saat',
  'Mekanizma', 'Durum',
  'Pop Sesi', 'Devam Edebildi', 'İlk 24s Şişlik',
  'Önceki ACL (aynı)', 'Önceki ACL (karşı)', 'Önceki Diz Yaralanması', 'Ailede ACL',
  'Yoğunluk Artışı', 'Önleyici Program',
  'Tegner (öncesi)', 'Tegner (şimdi)',
  'KOOS Belirti', 'KOOS Ağrı', 'KOOS GYA', 'KOOS Spor', 'KOOS YK', 'Marx',
  'KOOS ham', 'Marx ham',
  // ─── Hekim bölümü — hasta satırına sonradan yazılır ───
  'Vaka No', 'Hekim', 'Hekim Kayıt Zamanı',
  'Ekstansiyon Defisiti', 'Fleksiyon (°)', 'Efüzyon', 'Lachman', 'Pivot Shift',
  'Uyluk Atrofisi (cm)', 'MCL', 'PCL / PLK', 'Muayene Not',
  'MM Yırtık', 'MM Yeri', 'MM Tipi', 'LM Yırtık', 'LM Yeri', 'LM Tipi',
  'Kıkırdak Bölge', 'Kıkırdak ICRS', 'Kemik Kontüzyonu', 'Segond',
  'MR Diğer Bağ', 'MR Not',
  'Ameliyat Tarihi', 'Greft', 'Greft Çapı (mm)',
  'MM Girişim', 'LM Girişim', 'Ramp Tamiri', 'Ek Lateral Prosedür', 'Kıkırdak Girişimi',
  'İntra-op Not', 'Komplikasyon', 'Hekim Tamamlandı',
  // ─── Düzenleme izi ───
  'Veri Düzenleyen', 'Düzenleme Zamanı', 'Düzenlenen Alanlar',
];

// Vaka No hasta satırı yazılırken üretilir; hekim girişi bu anahtarla aynı satırı bulur.
const ACL_VAKA_COL = 47;            // 'Vaka No' sütunu (1 tabanlı)
const ACL_DOC_COL  = 48;            // hekim alanları buradan başlar
const ACL_DOC_LEN  = 34;            // 'Hekim' … 'Hekim Tamamlandı'

// Hastanın kendi doldurduğu bölümden SONRADAN düzeltilebilecek alanlar.
// KOOS / Marx / Tegner bilerek yok: onlar hastanın beyanı, hekim değiştiremez.
const ACL_EDIT_COLS = {
  patientId: 3, name: 4, age: 6, gender: 7, height: 8, weight: 9,
  dominant: 11, smoke: 12, injDate: 13, side: 15,
  sport: 16, level: 17, surface: 18, setting: 19,
  freqWeek: 20, sessionMin: 21, monthsYear: 22, yearsPlaying: 23,
  mechanism: 26, situation: 27, pop: 28, continued: 29, swelling: 30,
  priorAclSame: 31, priorAclOther: 32, priorKnee: 33, familyAcl: 34,
  rampUp: 35, preventionProg: 36,
};
const ACL_EDIT_NUM = ['age', 'height', 'weight', 'freqWeek', 'sessionMin', 'monthsYear', 'yearsPlaying'];
const ACL_AUDIT_COL = 82;           // 'Veri Düzenleyen' … 'Düzenlenen Alanlar'

function newVakaNo_() {
  return 'ACL-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}

function saveAclRow(data) {
  // Protokol no eşleştirme anahtarı; büyük/küçük harf farkı iki ayrı hasta yaratmasın
  data.patientId = String(data.patientId || '').trim().toUpperCase();
  const n = v => (v !== undefined && v !== null && v !== '') ? Number(v) : '';
  const sheet = ensureSheet(ACL_SHEET, ACL_HEADERS, [[4, 150], [5, 150], [16, 150], [17, 180], [45, 240]]);
  const ts = data.timestamp ? new Date(data.timestamp) : new Date();
  const hasta = [
    ts, Utilities.formatDate(ts, 'Europe/Istanbul', 'yyyy-MM-dd'),
    data.patientId || '', data.name || '', data.src || '',
    n(data.age), data.gender || '', n(data.height), n(data.weight), n(data.bmi),
    data.dominant || '', data.smoke || '',
    data.injDate || '', n(data.daysSinceInjury), data.side || '',
    data.sport || '', data.level || '', data.surface || '', data.setting || '',
    n(data.freqWeek), n(data.sessionMin), n(data.monthsYear), n(data.yearsPlaying),
    n(data.hoursWeek), n(data.hoursTotal),
    data.mechanism || '', data.situation || '',
    data.pop || '', data.continued || '', data.swelling || '',
    data.priorAclSame || '', data.priorAclOther || '', data.priorKnee || '', data.familyAcl || '',
    data.rampUp || '', data.preventionProg || '',
    n(data.tegnerPre), n(data.tegnerNow),
    n(data.koosS), n(data.koosP), n(data.koosA), n(data.koosSP), n(data.koosQ), n(data.marx),
    data.koosRaw || '', data.marxRaw || '',
  ];

  // Bir protokol no = bir ACL vakası. Aynı protokolle yeniden gelen anket
  // yeni satır açmaz, o satırın hasta bölümünü yeniden yazar — hekim verisi ve
  // Vaka No korunur. Hekimin anketsiz açtığı satır da böylece kendiliğinden dolar.
  // Karşı diz ayrı bir vakadır: taraf tutmuyorsa eşleştirme yapılmaz.
  const prot = String(data.patientId || '').trim().toUpperCase();
  const taraf = String(data.side || '').trim().toLowerCase();
  let row = 0, yeniden = false;
  const last = sheet.getLastRow();
  if (prot && last >= 2) {
    const v = sheet.getRange(2, 1, last - 1, ACL_HEADERS.length).getValues();
    for (let i = 0; i < v.length; i++) {
      if (String(v[i][2] || '').trim().toUpperCase() !== prot) continue;
      const rt = String(v[i][14] || '').trim().toLowerCase();
      if (rt && taraf && rt !== taraf) continue;      // karşı diz → ayrı vaka
      row = i + 2;
      yeniden = v[i][38] !== '' && v[i][38] !== null; // KOOS doluysa anket tekrar dolduruluyor
      break;
    }
  }
  if (row) {
    sheet.getRange(row, 1, 1, hasta.length).setValues([hasta]);
    if (yeniden) {
      // Hangi satırın üzerine yazıldığı kayda geçsin
      const iz = sheet.getRange(row, ACL_AUDIT_COL, 1, 3).getValues()[0];
      sheet.getRange(row, ACL_AUDIT_COL, 1, 3).setValues([[
        iz[0] || 'hasta',
        new Date(),
        ('Anket yeniden dolduruldu' + (iz[2] ? ' · ' + iz[2] : '')).slice(0, 2000),
      ]]);
    }
  } else {
    sheet.appendRow(hasta.concat([newVakaNo_()]).concat(new Array(ACL_DOC_LEN).fill('')));
    row = sheet.getLastRow();
  }
  colorGecenGun(sheet, row, 14, data.daysSinceInjury);
  [39, 40, 41, 42, 43].forEach(function (c) { colorKoos(sheet, row, c); });
}

// ─── Hekim girişi ────────────────────────────────────────────────────────────
// Hasta pre-op anketi doldurduğunda satır zaten açılmıştır. Hekim ameliyattan
// sonra müsait olduğunda aynı satırı Vaka No ile bulup kendi bölümünü yazar.
// Hasta hiç doldurmadıysa (vaka=='' ) hekim yeni bir satır açabilir.
function aclDocValues_(d) {
  const n = v => (v !== undefined && v !== null && v !== '') ? Number(v) : '';
  return [
    d.doctor || '', new Date(),
    d.extDeficit || '', n(d.flexion), d.effusion || '', d.lachman || '', d.pivot || '',
    n(d.atrophy), d.mcl || '', d.pcl || '', d.examNote || '',
    d.mmTear || '', d.mmSite || '', d.mmType || '',
    d.lmTear || '', d.lmSite || '', d.lmType || '',
    d.cartSite || '', d.cartIcrs || '', d.contusion || '', d.segond || '',
    d.otherLig || '', d.mrNote || '',
    d.surgDate || '', d.graft || '', n(d.graftDia),
    d.mmProc || '', d.lmProc || '', d.ramp || '', d.lateralProc || '', d.cartProc || '',
    d.opNote || '', d.complication || '', d.done ? 'Evet' : 'Hayır',
  ];
}

function saveAclDoctorRow(data) {
  data.patientId = String(data.patientId || '').trim().toUpperCase();
  const sheet = ensureSheet(ACL_SHEET, ACL_HEADERS, [[4, 150], [5, 150], [16, 150], [17, 180], [45, 240]]);
  const vals = aclDocValues_(data);
  let row = 0;
  const last = sheet.getLastRow();
  if (data.vaka && last >= 2) {
    const keys = sheet.getRange(2, ACL_VAKA_COL, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0] || '').trim() === String(data.vaka).trim()) { row = i + 2; break; }
    }
  }
  if (!row) {
    // Hasta anketi yok — hekim vakayı kendisi açıyor. Kimlik alanları doldurulur,
    // hasta sonradan doldurursa panelden bu satıra bağlanır.
    const ts = new Date();
    const bos = new Array(ACL_HEADERS.length).fill('');
    bos[0] = ts;
    bos[1] = Utilities.formatDate(ts, 'Europe/Istanbul', 'yyyy-MM-dd');
    bos[2] = data.patientId || '';
    bos[3] = data.name || '';
    bos[4] = data.doctor || '';
    bos[14] = data.side || '';
    bos[ACL_VAKA_COL - 1] = data.vaka || newVakaNo_();
    sheet.appendRow(bos);
    row = sheet.getLastRow();
  }
  sheet.getRange(row, ACL_DOC_COL, 1, ACL_DOC_LEN).setValues([vals]);
  sheet.getRange(row, ACL_DOC_COL + ACL_DOC_LEN - 1)
       .setBackground(data.done ? '#dcfce7' : '#fef9c3');
  return String(sheet.getRange(row, ACL_VAKA_COL).getValue() || '');
}

// QR ile gelen hastanın adı eksik/yanlış olabiliyor, hasta no hiç olmuyor,
// yaralanma tarihini yanlış hatırlayabiliyor. Hekim bu alanları düzeltebilsin.
// VKİ, geçen gün ve maruziyet saatleri girdilerden yeniden hesaplanır.
function saveAclPatientEdit(data) {
  if ('patientId' in data) data.patientId = String(data.patientId || '').trim().toUpperCase();
  const sheet = ensureSheet(ACL_SHEET, ACL_HEADERS, [[4, 150], [5, 150], [16, 150], [17, 180], [45, 240]]);
  const last = sheet.getLastRow();
  if (last < 2) throw new Error('kayıt yok');
  let row = 0;
  const keys = sheet.getRange(2, ACL_VAKA_COL, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0] || '').trim() === String(data.vaka || '').trim()) { row = i + 2; break; }
  }
  if (!row) throw new Error('vaka bulunamadı');

  const degisen = [];
  Object.keys(ACL_EDIT_COLS).forEach(function (k) {
    if (!(k in data)) return;                       // gönderilmeyen alana dokunma
    const col = ACL_EDIT_COLS[k];
    let v = data[k];
    if (ACL_EDIT_NUM.indexOf(k) >= 0) v = (v === '' || v === null) ? '' : Number(v);
    const eski = sheet.getRange(row, col).getValue();
    if (String(eski) !== String(v)) {
      sheet.getRange(row, col).setValue(v);
      degisen.push(ACL_HEADERS[col - 1]);
    }
  });

  const r = sheet.getRange(row, 1, 1, ACL_HEADERS.length).getValues()[0];
  const sayi = c => { const x = Number(r[c - 1]); return isNaN(x) ? 0 : x; };

  // VKİ
  const boy = sayi(8), kilo = sayi(9);
  if (boy > 0 && kilo > 0) sheet.getRange(row, 10).setValue(Math.round(kilo / Math.pow(boy / 100, 2) * 10) / 10);

  // Geçen gün: yaralanmadan anketin doldurulduğu güne
  const yar = r[12], anket = r[1];
  if (yar && anket) {
    const d1 = new Date(yar), d2 = new Date(anket);
    if (!isNaN(d1) && !isNaN(d2)) {
      const g = Math.round((d2 - d1) / 86400000);
      sheet.getRange(row, 14).setValue(g);
      colorGecenGun(sheet, row, 14, g);
    }
  }

  // Maruziyet: haftada kez × seans dk ÷ 60 ; toplam = haftalık × 4,33 × ay × yıl
  const kez = sayi(20), dk = sayi(21), ay = sayi(22), yil = sayi(23);
  if (kez > 0 && dk > 0) {
    const haftalik = Math.round(kez * dk / 60 * 10) / 10;
    sheet.getRange(row, 24).setValue(haftalik);
    if (ay > 0 && yil > 0) sheet.getRange(row, 25).setValue(Math.round(haftalik * 4.33 * ay * yil));
  }

  sheet.getRange(row, ACL_AUDIT_COL, 1, 3)
       .setValues([[data.editor || '', new Date(), degisen.join(', ')]]);
  return degisen;
}

// Telefonda hızlı açılsın diye tam getData yerine hafif vaka listesi.
// Vaka No'su olmayan eski satırlara bu sırada anahtar yazılır.
function aclCaseList() {
  const sheet = ensureSheet(ACL_SHEET, ACL_HEADERS, [[4, 150], [5, 150], [16, 150], [17, 180], [45, 240]]);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const d = sheet.getRange(2, 1, last - 1, ACL_HEADERS.length).getValues();
  const out = [];
  for (let i = 0; i < d.length; i++) {
    const r = d[i];
    if (!String(r[2] || '').trim() && !String(r[3] || '').trim()) continue;
    let vaka = String(r[ACL_VAKA_COL - 1] || '').trim();
    if (!vaka) { vaka = newVakaNo_(); sheet.getRange(i + 2, ACL_VAKA_COL).setValue(vaka); }
    out.push({
      vaka: vaka,
      ad: String(r[3] || ''), id: String(r[2] || ''),
      tarih: r[1] instanceof Date ? Utilities.formatDate(r[1], 'Europe/Istanbul', 'yyyy-MM-dd') : String(r[1] || ''),
      yaralanma: r[12] instanceof Date ? Utilities.formatDate(r[12], 'Europe/Istanbul', 'yyyy-MM-dd') : String(r[12] || ''),
      gun: r[13] === '' ? null : Number(r[13]),
      taraf: String(r[14] || ''), spor: String(r[15] || ''),
      yas: r[5] === '' ? null : Number(r[5]),
      hastaDoldu: String(r[3] || '').trim() !== '' && r[38] !== '',
      hekim: String(r[ACL_DOC_COL - 1] || ''),
      amelTarih: r[ACL_DOC_COL + 22] instanceof Date
        ? Utilities.formatDate(r[ACL_DOC_COL + 22], 'Europe/Istanbul', 'yyyy-MM-dd')
        : String(r[ACL_DOC_COL + 22] || ''),
      tamam: String(r[ACL_DOC_COL + ACL_DOC_LEN - 2] || '') === 'Evet',
    });
  }
  return out.reverse();   // en yeni başa
}

function aclCaseDetail(vaka) {
  const sheet = ensureSheet(ACL_SHEET, ACL_HEADERS, [[4, 150], [5, 150], [16, 150], [17, 180], [45, 240]]);
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const keys = sheet.getRange(2, ACL_VAKA_COL, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0] || '').trim() === String(vaka).trim()) {
      const r = sheet.getRange(i + 2, 1, 1, ACL_HEADERS.length).getValues()[0];
      const o = {};
      for (let c = 0; c < ACL_HEADERS.length; c++) {
        const v = r[c];
        o[ACL_HEADERS[c]] = v instanceof Date
          ? Utilities.formatDate(v, 'Europe/Istanbul', 'yyyy-MM-dd') : v;
      }
      return o;
    }
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════════════════
// ROTATOR MANŞET — AMELİYAT ÖNCESİ DEĞERLENDİRME
// Ölçekler: ASES (0-100) + WORC (0-100), ikisi de yüksek = iyi.
// ACL ile aynı mimari: bir protokol no = bir vaka, satır Vaka No ile anahtarlanır.
// ═══════════════════════════════════════════════════════════════════════════
const RC_SHEET = 'PROM_RC_PreOp';
const RC_HEADERS = [
  'Zaman Damgası', 'Tarih', 'Hasta ID', 'Ad Soyad', 'Kaynak',
  'Yaş', 'Cinsiyet', 'Boy (cm)', 'Kilo (kg)', 'VKİ',
  'Dominant El', 'Taraf', 'Dominant Omuz mu', 'Sigara', 'Ek Hastalık',
  'Şikâyet Süresi (ay)', 'Başlangıç', 'Travma Tarihi', 'Travmadan Geçen Gün',
  'İş Kazası / Meslek Hastalığı', 'İş Yükü',
  'Gece Ağrısı', 'Üzerine Yatamıyor', 'İstirahat Ağrısı', 'Haftada Kaç Gece',
  'Fizik Tedavi', 'Enjeksiyon Sayısı', 'Son Enjeksiyon',
  'Önceki Ameliyat (aynı omuz)', 'Karşı Omuz', 'Boyun Sorunu',
  'ASES Ağrı (0-10)', 'ASES',
  'WORC Belirti', 'WORC Spor', 'WORC İş', 'WORC Yaşam', 'WORC Duygu', 'WORC',
  'Beklenti', 'Beklenen Düzelme (%)',
  'ASES ham', 'WORC ham', 'Vaka No',
  // ─── Sigara ayrıntısı, spor ve uyku ───
  'Sigara (adet/gün)', 'Sigara (yıl)', 'Paket-Yıl', 'Bıraktı (yıl önce)',
  'Spor Yapıyor', 'Spor', 'Spor Düzeyi', 'Spor Haftada Kez', 'Spor Kaç Yıldır',
  'Baş Üstü Kol Kullanımı',
  'Uyku Pozisyonu', 'Pozisyon No', 'Pozisyon Tipi', 'Yattığı Kol',
  'Yatak Partneri', 'Yatağın Tarafı', 'Ağrı Saat Dilimi',
];
const RC_VAKA_COL = 44;

function saveRcRow(data) {
  data.patientId = String(data.patientId || '').trim().toUpperCase();
  const n = v => (v !== undefined && v !== null && v !== '') ? Number(v) : '';
  const sheet = ensureSheet(RC_SHEET, RC_HEADERS, [[4, 150], [5, 150], [15, 190], [17, 190], [43, 240], [44, 300]]);
  const ts = data.timestamp ? new Date(data.timestamp) : new Date();
  const hasta = [
    ts, Utilities.formatDate(ts, 'Europe/Istanbul', 'yyyy-MM-dd'),
    data.patientId || '', data.name || '', data.src || '',
    n(data.age), data.gender || '', n(data.height), n(data.weight), n(data.bmi),
    data.dominant || '', data.side || '', data.dominantSide || '',
    data.smoke || '', data.comorbid || '',
    n(data.durationMonths), data.onset || '',
    data.traumaDate || '', n(data.daysSinceTrauma),
    data.workComp || '', data.workLoad || '',
    data.nightPain || '', data.sideLie || '', data.restPain || '', n(data.nightsPerWeek),
    data.physio || '', data.injections || '', data.injectionTiming || '',
    data.prevSurgSame || '', data.prevOther || '', data.neckProb || '',
    n(data.asesPain), n(data.ases),
    n(data.worcS), n(data.worcSp), n(data.worcW), n(data.worcL), n(data.worcE), n(data.worc),
    data.expectGoal || '', n(data.expectPct),
    data.asesRaw || '', data.worcRaw || '',
  ];
  const ek = [
    n(data.smokePerDay), n(data.smokeYears), n(data.packYears), n(data.smokeQuitYears),
    data.sportYes || '', data.sport || '', data.sportLevel || '',
    n(data.sportWeek), n(data.sportYears), data.overheadSport || '',
    data.sleepPos || '', n(data.sleepPosNo), data.sleepPosBase || '',
    data.lieOnArm || '', data.partner || '', data.bedSide || '', data.painHour || '',
  ];

  // Bir protokol no + bir taraf = bir vaka. Aynı omuz için gelen ikinci anket
  // yeni satır açmaz, satırı yeniden yazar. Karşı omuz ayrı vakadır.
  const prot = data.patientId, taraf = String(data.side || '').trim().toLowerCase();
  let row = 0;
  const last = sheet.getLastRow();
  if (prot && last >= 2) {
    const v = sheet.getRange(2, 1, last - 1, RC_HEADERS.length).getValues();
    for (let i = 0; i < v.length; i++) {
      if (String(v[i][2] || '').trim().toUpperCase() !== prot) continue;
      const rt = String(v[i][11] || '').trim().toLowerCase();
      if (rt && taraf && rt !== taraf) continue;
      row = i + 2; break;
    }
  }
  if (row) {
    sheet.getRange(row, 1, 1, hasta.length).setValues([hasta]);
    sheet.getRange(row, hasta.length + 2, 1, ek.length).setValues([ek]);   // Vaka No atlanır
  } else {
    sheet.appendRow(hasta
      .concat(['RC-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase()])
      .concat(ek));
    row = sheet.getLastRow();
  }

  colorSure(sheet, row, 16, data.durationMonths);
  colorEnj(sheet, row, 27, data.injections, data.injectionTiming);
  [33, 34, 35, 36, 37, 38, 39].forEach(function (c) { colorKoos(sheet, row, c); });   // ASES + 6 WORC
}

// Şikâyet süresi uzadıkça onarım ve iyileşme şansı düşüyor
function colorSure(sheet, row, col, ay) {
  const a = Number(ay);
  if (!a && a !== 0) { sheet.getRange(row, col).setBackground(null); return; }
  sheet.getRange(row, col).setBackground(a <= 6 ? '#dcfce7' : a <= 12 ? '#fef9c3' : '#fee2e2');
}

// Çok sayıda veya son 6 ay içindeki kortikosteroid enjeksiyonu retear riskini artırıyor
function colorEnj(sheet, row, col, sayi, zaman) {
  const s = String(sayi || ''), z = String(zaman || '');
  if (!s || s === 'Hiç olmadım') { sheet.getRange(row, col).setBackground(null); return; }
  const cok = s === '2 kez' || s === '3 ve üzeri';
  const yakin = z === 'Son 3 ay içinde' || z === '3–6 ay önce';
  sheet.getRange(row, col).setBackground(cok && yakin ? '#fee2e2' : (cok || yakin) ? '#fef9c3' : null);
}

// Yaralanma-ameliyat arası süre: 6 ay üstü menisküs riski belirgin artıyor
function colorGecenGun(sheet, row, col, v) {
  const g = Number(v);
  if (!g && g !== 0) { sheet.getRange(row, col).setBackground(null); return; }
  sheet.getRange(row, col).setBackground(g <= 90 ? '#dcfce7' : g <= 180 ? '#fef9c3' : '#fee2e2');
}

function colorKoos(sheet, row, col) {
  const v = Number(sheet.getRange(row, col).getValue());
  if (!v && v !== 0) return;
  sheet.getRange(row, col).setBackground(v >= 75 ? '#dcfce7' : v >= 50 ? '#fef9c3' : '#fee2e2');
}
