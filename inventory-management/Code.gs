const SCRIPT_PROP = PropertiesService.getScriptProperties();
const SHEET_NAME = 'Inventory';
// Columns mapping:
// A: วดป. (Date/Time)
// B: หมายเลขงาน (Job Number)
// C: ชื่องานก่อสร้าง (Construction Project Name)
// D: หมายเลขเสา (Pole Number)
// E: รับ (Stock In)
// F: จ่าย (Stock Out)
// G: คงเหลือ (Balance)
// H: ผู้เบิก / ผู้ส่งคืน (Requester/Returner Name)

function setup() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty("key", doc.getId());
}

function doGet(e) {
  return handleResponse(e);
}

function doPost(e) {
  return handleResponse(e);
}

function handleResponse(e) {
  // Use a public lock to prevent concurrent writing issues
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = doc.getSheetByName(SHEET_NAME);
    
    // Create sheet if it does not exist
    if (!sheet) {
      sheet = doc.insertSheet(SHEET_NAME);
      const headers = ['วดป.', 'หมายเลขงาน', 'ชื่องานก่อสร้าง', 'หมายเลขเสา', 'รับ', 'จ่าย', 'คงเหลือ', 'ผู้เบิก / ผู้ส่งคืน'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    } else if (sheet.getLastRow() === 0) {
      const headers = ['วดป.', 'หมายเลขงาน', 'ชื่องานก่อสร้าง', 'หมายเลขเสา', 'รับ', 'จ่าย', 'คงเหลือ', 'ผู้เบิก / ผู้ส่งคืน'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    // Check if it's a GET request to fetch data
    if (e.parameter && e.parameter.action === 'getInventory') {
      return getInventory(sheet);
    }

    // Default: Handle POST submission
    const data = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter;
    
    const poleId = data['หมายเลขเสา'];
    const receive = parseFloat(data['รับ']) || 0;
    const dispense = parseFloat(data['จ่าย']) || 0;
    
    // Calculate new balance based on previous records for the same 'หมายเลขเสา'
    const previousBalance = getLatestBalance(sheet, poleId);
    const newBalance = previousBalance + receive - dispense;
    
    // Construct new row data
    const newRow = [
      data['วดป.'] || new Date().toISOString(),
      data['หมายเลขงาน'],
      data['ชื่องานก่อสร้าง'],
      poleId,
      receive > 0 ? receive : "", // keep 0 as blank for readability
      dispense > 0 ? dispense : "",
      newBalance,
      data['ผู้เบิก / ผู้ส่งคืน']
    ];
    
    sheet.appendRow(newRow);
    
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'success', 
        balance: newBalance,
        message: 'Record added successfully'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Searches from the bottom of the sheet to find the last known balance for a given Pole ID.
 */
function getLatestBalance(sheet, poleId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0; // No data rows
  
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); 
  // We only need up to column G (index 6)
  // Col A=0, B=1, C=2, D=3 (Pole ID), E=4, F=5, G=6 (Balance)
  
  let latestBalance = 0;
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][3]).trim() === String(poleId).trim()) {
      latestBalance = parseFloat(data[i][6]) || 0;
      break;
    }
  }
  return latestBalance;
}

/**
 * Returns all inventory history records 
 */
function getInventory(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', data: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = data[0];
  const items = data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  
  // CORS wrapper is required if fetching from frontend
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'success', data: items.reverse() })) // Newest first
    .setMimeType(ContentService.MimeType.JSON);
}

