const SHEET_KUPON = 'Kupon';
const SHEET_LOG = 'Log';
const SHEET_CONFIG = 'Konfigurasi';
const SHEET_QR = 'QR Codes';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🐄 Qurban Scanner')
    .addItem('1. Inisialisasi Sheet', 'initSheet')
    .addItem('2. Generate QR Codes', 'generateQRCodes')
    .addItem('Lihat Statistik', 'showStatsPrompt')
    .addToUi();
}

function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetKupon = ss.getSheetByName(SHEET_KUPON) || ss.insertSheet(SHEET_KUPON);
  sheetKupon.getRange('A1:D1').setValues([['Kode Kupon', 'Status', 'Waktu Scan', 'Scanner ID']]).setFontWeight('bold');
  sheetKupon.setFrozenRows(1);

  let sheetLog = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  sheetLog.getRange('A1:D1').setValues([['Timestamp', 'Kode Kupon', 'Aksi', 'Scanner ID']]).setFontWeight('bold');
  sheetLog.setFrozenRows(1);

  let sheetConfig = ss.getSheetByName(SHEET_CONFIG) || ss.insertSheet(SHEET_CONFIG);
  sheetConfig.getRange('A1:B3').setValues([
    ['Pengaturan', 'Nilai'], 
    ['prefix', 'QRB-'],
    ['admin_pin', '123456']
  ]).setFontWeight('bold');
  sheetConfig.setFrozenRows(1);
}

function handleBulkAdd(jumlah) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetKupon = ss.getSheetByName(SHEET_KUPON);
  const sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  
  if (!sheetKupon || !sheetConfig) return { success: false, message: 'Sheet belum diinisialisasi.' };

  const prefix = sheetConfig.getRange('B2').getValue() || 'QRB-';
  const lastRow = sheetKupon.getLastRow();
  const startNumber = lastRow > 1 ? parseInt(sheetKupon.getRange(lastRow, 1).getValue().replace(prefix, '')) + 1 : 1;
  
  let newData = [];
  for (let i = 0; i < jumlah; i++) {
    const code = prefix + String(startNumber + i).padStart(7, '0');
    newData.push([code, 'Tersedia', '', '']);
  }
  
  sheetKupon.getRange(lastRow + 1, 1, newData.length, 4).setValues(newData);
  return { success: true, message: `${jumlah} kupon berhasil di-generate.` };
}

function generateQRCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetKupon = ss.getSheetByName(SHEET_KUPON);
  if (!sheetKupon) return;

  let sheetQR = ss.getSheetByName(SHEET_QR) || ss.insertSheet(SHEET_QR);
  sheetQR.clear();

  const lastRow = sheetKupon.getLastRow();
  if (lastRow < 2) return;

  const data = sheetKupon.getRange(2, 1, lastRow - 1, 1).getValues().filter(row => row[0] !== '');
  
  const kolomPerBaris = 7; 
  let gridData = [];
  let currentRowQR = [];
  let currentRowText = [];

  for (let i = 0; i < data.length; i++) {
    const kode = data[i][0];
    const urlQR = `=IMAGE("https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(kode)}&size=150x150")`;
    
    currentRowQR.push(urlQR);
    currentRowText.push(kode);
    
    if (currentRowQR.length === kolomPerBaris || i === data.length - 1) {
      while (currentRowQR.length < kolomPerBaris) {
        currentRowQR.push("");
        currentRowText.push("");
      }
      
      gridData.push(currentRowQR);
      gridData.push(currentRowText);
      
      currentRowQR = [];
      currentRowText = [];
    }
  }

  if (gridData.length > 0) {
    sheetQR.getRange(1, 1, gridData.length, kolomPerBaris).setValues(gridData);
    
    const lebarKolom = 100;
    const tinggiQR = 100;
    const tinggiTeks = 25;

    for (let j = 1; j <= kolomPerBaris; j++) {
      sheetQR.setColumnWidth(j, lebarKolom);
    }

    for (let i = 1; i <= gridData.length; i++) {
      sheetQR.setRowHeight(i, i % 2 !== 0 ? tinggiQR : tinggiTeks);
    }
    
    sheetQR.getDataRange()
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontSize(8);
  }
}

function showStatsPrompt() {
  const stats = getStats();
  SpreadsheetApp.getUi().alert(`Total: ${stats.total}\nDigunakan: ${stats.digunakan}\nSisa: ${stats.sisa}`);
}

function getStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KUPON);
  if (!sheet) return { total: 0, digunakan: 0, sisa: 0 };
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0, digunakan: 0, sisa: 0 };
  
  const statuses = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  const total = statuses.length;
  const digunakan = statuses.filter(s => s === 'Digunakan').length;
  
  return { total: total, digunakan: digunakan, sisa: total - digunakan };
}

function processScan(code, scannerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetKupon = ss.getSheetByName(SHEET_KUPON);
  const sheetLog = ss.getSheetByName(SHEET_LOG);
  
  const dataRange = sheetKupon.getRange(2, 1, Math.max(1, sheetKupon.getLastRow() - 1), 2);
  const data = dataRange.getValues();
  const timestamp = new Date().toLocaleString('id-ID');
  
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === code) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    sheetLog.appendRow([timestamp, code, 'TIDAK TERDAFTAR', scannerId]);
    return { success: false, status: 'TIDAK TERDAFTAR', message: 'Kupon palsu / tidak ditemukan.' };
  }

  const currentStatus = data[rowIndex][1];
  if (currentStatus === 'Digunakan') {
    sheetLog.appendRow([timestamp, code, 'DITOLAK (SUDAH DIGUNAKAN)', scannerId]);
    return { success: false, status: 'SUDAH DIGUNAKAN', message: 'Kupon ini sudah ditukarkan sebelumnya.' };
  }

  sheetKupon.getRange(rowIndex + 2, 2, 1, 3).setValues([['Digunakan', timestamp, scannerId]]);
  
  sheetKupon.getRange(rowIndex + 2, 1, 1, 4).setBackground('#d9ead3');
  
  sheetLog.appendRow([timestamp, code, 'KUPON VALID', scannerId]);
  return { success: true, status: 'KUPON VALID', message: 'Kupon sah, berikan daging.' };
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result = {};
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetConfig = ss.getSheetByName(SHEET_CONFIG);
    
    let serverPin = '123456'; 
    if (sheetConfig) {
      const pinValue = sheetConfig.getRange('B3').getValue();
      if (pinValue) serverPin = String(pinValue);
    }

    const adminActions = ['init', 'bulk', 'manual'];

    if (adminActions.includes(action)) {
      if (String(payload.adminPin) !== serverPin) {
         return ContentService.createTextOutput(JSON.stringify({ 
           success: false, 
           message: 'Akses Ditolak: PIN Admin tidak valid.' 
         })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    switch (action) {
      case 'scan':
        result = processScan(payload.code, payload.scannerId);
        break;
      case 'stats':
        result = getStats();
        break;
      case 'init':
        initSheet();
        result = { success: true, message: 'Inisialisasi selesai' };
        break;
      case 'bulk':
        result = handleBulkAdd(parseInt(payload.jumlah) || 0);
        break;
      case 'manual':
        const sheetKupon = ss.getSheetByName(SHEET_KUPON);
        sheetKupon.appendRow([payload.code, 'Tersedia', '', '']);
        result = { success: true, message: `Kupon ${payload.code} ditambahkan.` };
        break;
      default:
        result = { success: false, message: 'Aksi tidak dikenal' };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Backend Aktif").setMimeType(ContentService.MimeType.TEXT);
}