// ============================================================
// PROM Takip Sistemi – Google Apps Script Backend
// Omuz Kalsifik Tendinit | VAS · ROM · DASH
// Bu kodu Google Apps Script'e yapıştırın ve deploy edin.
// ============================================================

const SHEET_NAME = 'PROM_Kayitlar';

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
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Sütun sırası (0-based index):
  // 0:timestamp  1:date      2:time       3:patientId  4:patientName  5:diagnosis
  // 6:surgeryDate 7:daysPostOp 8:surgeryLeg
  // 9:vas  10:rom_abduction  11:rom_flexion  12:rom_ir  13:rom_ir_label  14:rom_er
  // 15:dash  16:constant  17:sleep_vas  18:sf12_pcs  19:sf12_mcs  20:followup
  const headers = [
    'Zaman Damgası', 'Tarih', 'Saat', 'Hasta ID', 'Hasta Adı', 'Tanı',
    'Girişim Tarihi', 'Post-op Gün', 'Taraf',
    'VAS (0-10)', 'Abduksiyon (°)', 'Öne Fleksiyon (°)', 'İç Rotasyon', 'İR Pozisyon', 'Dış Rotasyon (°)',
    'DASH (0-100)', 'Constant (0-100)', 'Uyku VAS (0-10)', 'SF-12 PCS', 'SF-12 MCS', 'Takip',
  ];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);  // Zaman Damgası
    sheet.setColumnWidth(5, 160);  // Hasta Adı
  }

  // Başlık satırını yaz / yeni sütun eklendiğinde mevcut sayfada tamamla
  if (sheet.getRange(1, headers.length).getValue() !== headers[headers.length - 1]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const hr = sheet.getRange(1, 1, 1, headers.length);
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

  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  // Sütun numaraları (1-based): VAS=10, Abd=11, Flex=12, ER=15, DASH=16, Constant=17, Uyku=18, PCS=19, MCS=20
  colorVAS(sheet, lastRow, 10, data.vas);
  colorROM(sheet, lastRow, 11, data.rom_abduction, 150, 90);
  colorROM(sheet, lastRow, 12, data.rom_flexion,   150, 90);
  colorROM(sheet, lastRow, 15, data.rom_er,          60, 30);
  colorDASH(sheet, lastRow, 16, data.dash);
  colorConstant(sheet, lastRow, 17, data.constant);
  colorSleep(sheet, lastRow, 18, data.sleep_vas);
  colorSF12(sheet, lastRow, 19, data.sf12_pcs);
  colorSF12(sheet, lastRow, 20, data.sf12_mcs);
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
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ rows: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const data = sheet.getDataRange().getValues();
  return ContentService
    .createTextOutput(JSON.stringify({ rows: data }))
    .setMimeType(ContentService.MimeType.JSON);
}
