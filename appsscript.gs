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

function sheetNameFor(diag) {
  return DIAG_SHEETS[diag] || DEFAULT_SHEET;
}

// Takip dönemi sırası — aynı hastanın satırları tabloda bu sıraya göre alt alta dizilir
const FU_ORDER = { 'Pre-op': 0, '2. Hafta': 1, '6. Hafta': 2, '6. Ay': 3 };
function fuRank(label) {
  return FU_ORDER.hasOwnProperty(label) ? FU_ORDER[label] : 99;
}

// Yeni kaydın gideceği satır numarası: aynı hastanın bloğu içinde dönem sırasına
// göre konum bulur (hasta anahtarı: Hasta ID, yoksa ad). Hasta yoksa null → sona eklenir.
function findInsertRow(sheet, data) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
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
    if (insertAt === -1 && fuRank(String(values[i][20] || '')) > newRank) insertAt = i;
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
  return ContentService
    .createTextOutput(JSON.stringify({ rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}
