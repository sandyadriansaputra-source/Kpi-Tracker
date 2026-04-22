/**
 * PT PEGADAIAN CP WONOMULYO - KPI API
 * Update: Full Script with Column O Fixed, Optimized Range & Auto Reset
 * Support: Harian, Bulanan, Tahunan
 */

function doGet(e) {
  const sheetName = e.parameter.sheet || "KPI TRACKER THN";
  
  if (e.parameter.action === "read") {
    const data = getKpiData(sheetName);
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutput("API KPI Pegadaian Aktif - Status: OK");
}

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  let result;

  if (params.action === "reset_all") {
    result = manualResetSheet(params.sheet);
  } 
  else {
    result = saveUpdate(params.rowIdx, params.value, params.sheet);
  }

  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
}

function getKpiData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName); 
  if (!sheet) return { items: [], date: "", infoM2: "" };

  const dataDate = sheet.getRange("J2").getDisplayValue();
  const dataM2 = sheet.getRange("M2").getDisplayValue();

  const lastRow = sheet.getLastRow();
  // OPTIMASI: Hanya ambil data sampai kolom O (15) untuk mempercepat loading
  const data = sheet.getRange(1, 1, lastRow, 15).getValues();
  
  const kpiItems = data.slice(1).filter(row => row[2] !== "").map((row, index) => {
    let rawAch = parseFloat(row[8]) || 0; 
    let formattedAch = rawAch <= 1.1 ? (rawAch * 100).toFixed(1) : rawAch.toFixed(1);
    
    return {
      rowIdx: index + 2,
      kategori: String(row[0] || "").trim(),
      unit: String(row[1] || "").trim(),
      komponen: String(row[2] || "").trim(),
      bobot: row[3] || 0,
      target: parseFloat(row[4]) || 0,
      targetBln: parseFloat(row[4]) || 0,
      targetThn: parseFloat(row[12]) || 0,
      saldoAwal: parseFloat(row[5]) || 0,
      realisasi: parseFloat(row[11]) || 0, // Kolom L
      persentase: formattedAch,
      nilaiKpi: parseFloat(row[10]) || 0,
      
      // Keterangan di Kolom N (Index 13)
      keterangan: String(row[13] || "").trim()
    };
  });

  return { items: kpiItems, date: dataDate, infoM2: dataM2 };
}

function saveUpdate(rowIndex, value, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const inputVal = parseFloat(value) || 0;

  // Langsung isi ke Kolom O (Kolom 15)
  const targetCol = 15;
  const rangeUpdate = sheet.getRange(rowIndex, targetCol);
  const currentVal = parseFloat(rangeUpdate.getValue()) || 0;
  rangeUpdate.setValue(currentVal + inputVal);

  // Set formula di kolom L (12): Saldo Awal (F) + Update (O)
  sheet.getRange(rowIndex, 12).setFormula("=F" + rowIndex + "+O" + rowIndex);
  
  return "Berhasil Update!";
}

/**
 * FUNGSI RESET: Membersihkan hanya Kolom O
 */
function manualResetSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Sheet tidak ditemukan";

  const colO = 15; // Kolom O
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    // Kosongkan data di Kolom O
    sheet.getRange(2, colO, lastRow - 1).clearContent();
    
    // Kembalikan formula kolom L agar hanya merujuk ke saldo awal
    const rangeL = sheet.getRange(2, 12, lastRow - 1);
    rangeL.setFormula("=F2"); 
    
    return "Reset Kolom O Berhasil!";
  }
  return "Data sudah kosong";
}

/**
 * FUNGSI AUTO RESET: Dijalankan otomatis oleh trigger harian
 */
function autoResetJob() {
  const targetSheets = ["KPI TRACKER THN", "KPI TRACKER BLN", "UPDATE DATA HARIAN"]; 
  targetSheets.forEach(sheetName => {
    manualResetSheet(sheetName);
  });
}

/**
 * FUNGSI TRIGGER: Jalankan INI SEKALI SAJA di editor untuk mengaktifkan jadwal
 */
function createDailyTrigger() {
  // Hapus trigger lama jika sudah ada agar tidak duplikat
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'autoResetJob') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }

  // Buat trigger baru setiap jam 00:00 WITA (GMT+8)
  ScriptApp.newTrigger('autoResetJob')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .inTimezone("GMT+8") 
    .create();
}
