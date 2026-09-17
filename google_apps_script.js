// --------------------------------------------------------------------------------
// OSG myG PORTAL - BACKEND SCRIPT (v6 - Robust row matching + SR No fix)
// --------------------------------------------------------------------------------
const SHEET_NAME = "myG-OSG CUSTOMER COMPLAINT DATA";
const HEADER_ROW_INDEX = 1;
const CLAIM_ID_COLUMN  = "Claim ID";
const SR_NO_COLUMN     = "SR No";

// Columns that should NEVER be auto-added to the Google Sheet.
const SHEET_BLACKLIST = [
  "Follow Up - Notes", "Assigned Staff", "Follow Up - Dates",
  "Claim Settled Date", "Repair Feedback Completed (Yes/No)",
  "Customer Confirmation", "Approval Mail Received From Onsitego (Yes/No)",
  "Mail Sent To Store (Yes/No)", "Invoice Generated (Yes/No)",
  "Invoice Sent To Onsitego (Yes/No)", "Settlement Mail to Accounts(Yes/No)",
  "Settled With Accounts (Yes/No)", "Complete (Yes/No)",
  "Last Updated Timestamp", "Last_Notified_Status",
  "Replacement: Confirmation Pending", "Replacement: OSG Approval",
  "Replacement: Mail to Store", "Replacement: Invoice Generated",
  "Replacement: Invoice Sent to OSG", "Replacement: Settled with Accounts",
  "Replacement: Settlement Mail to Accounts",
  "Approval Mail Received Date", "Mail Sent To Store Date",
  "Invoice Generated Date", "Invoice Sent To Onsitego Date",
  "Settlement Mail to Accounts Date", "Feedback Rating",
  "Settled Time (TAT)", "Settled With Accounts (Yes/No)",
  "Complete", "claim_id", "sr_no", "_legacy_sr_lookup",
  "Mobile Number", "Date"
];

// -----------------------------------------
function _getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([SR_NO_COLUMN, "Submitted Date", "Customer Name", "Mobile",
      "Branch", "Product", "Issue", "Status", CLAIM_ID_COLUMN]);
  }
  return sheet;
}

function _readHeaders(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const raw = sheet.getRange(HEADER_ROW_INDEX, 1, 1, lastCol).getValues()[0] || [];
  return raw.map(h => (h === null || h === undefined) ? "" : String(h).trim());
}

function _colIndex(headers, exactName) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === exactName) return i;
  }
  return -1;
}

function _ensureColumn(sheet, headers, exactName) {
  let idx = _colIndex(headers, exactName);
  if (idx === -1) {
    const newCol = headers.length + 1;
    sheet.getRange(HEADER_ROW_INDEX, newCol).setValue(exactName);
    headers.push(exactName);
    idx = headers.length - 1;
  }
  return idx;
}

function _ensureBodyColumns(sheet, headers, body) {
  for (const key in body) {
    if (!body.hasOwnProperty(key)) continue;
    if (SHEET_BLACKLIST.indexOf(key) !== -1) continue;
    if (_colIndex(headers, key) === -1) {
      sheet.getRange(HEADER_ROW_INDEX, headers.length + 1).setValue(key);
      headers.push(key);
    }
  }
}

function _currentIst() {
  return Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
}

/**
 * ROBUST row matching:
 * 1. Exact match in Claim ID column
 * 2. Exact match in SR No column  
 * 3. Legacy CLM lookup in SR No column
 * 4. BRUTE FORCE: scan every cell in every row for the CLM value
 *    (handles any column arrangement, old rows, etc.)
 */
function _findRowIndex(sheet, headers, incomingClaimId, incomingSrNo, legacyLookup) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_INDEX) return -1;

  const dataRows  = lastRow - HEADER_ROW_INDEX;
  const numCols   = headers.length;

  // Read ALL row data at once (one API call)
  const allData = sheet.getRange(HEADER_ROW_INDEX + 1, 1, dataRows, numCols).getValues();

  const claimIdColIdx = _colIndex(headers, CLAIM_ID_COLUMN);
  const srNoColIdx    = _colIndex(headers, SR_NO_COLUMN);

  for (let i = 0; i < dataRows; i++) {
    const row = allData[i].map(String);

    // 1. Claim ID column match (most reliable for new rows)
    if (incomingClaimId && claimIdColIdx !== -1 && row[claimIdColIdx] === incomingClaimId) {
      Logger.log("Found row " + (HEADER_ROW_INDEX + 1 + i) + " via Claim ID column");
      return HEADER_ROW_INDEX + 1 + i;
    }
    // 2. SR No column match (real SR numbers)
    if (incomingSrNo && srNoColIdx !== -1 && row[srNoColIdx] === incomingSrNo) {
      Logger.log("Found row " + (HEADER_ROW_INDEX + 1 + i) + " via SR No match");
      return HEADER_ROW_INDEX + 1 + i;
    }
    // 3. Legacy CLM value stored in SR No column
    if (legacyLookup && srNoColIdx !== -1 && row[srNoColIdx] === legacyLookup) {
      Logger.log("Found row " + (HEADER_ROW_INDEX + 1 + i) + " via legacy SR No lookup");
      return HEADER_ROW_INDEX + 1 + i;
    }
    // 4. BRUTE FORCE: search every cell in this row for the CLM value
    if (incomingClaimId) {
      for (let c = 0; c < row.length; c++) {
        if (row[c] === incomingClaimId) {
          Logger.log("Found row " + (HEADER_ROW_INDEX + 1 + i) + " via brute-force scan (col " + c + ")");
          return HEADER_ROW_INDEX + 1 + i;
        }
      }
    }
  }
  Logger.log("Row NOT found for Claim ID: " + incomingClaimId + " / SR No: " + incomingSrNo);
  return -1;
}

// -----------------------------------------
// POST handler
// -----------------------------------------
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      body = e.parameter;
    }

    Logger.log("Received payload: " + JSON.stringify(body));

    const sheet = _getSheet();
    var headers = _readHeaders(sheet);

    _ensureColumn(sheet, headers, CLAIM_ID_COLUMN);
    _ensureColumn(sheet, headers, SR_NO_COLUMN);
    _ensureBodyColumns(sheet, headers, body);
    headers = _readHeaders(sheet);

    const claimIdColIdx    = _colIndex(headers, CLAIM_ID_COLUMN);
    const srNoColIdx       = _colIndex(headers, SR_NO_COLUMN);
    const submittedDateIdx = _colIndex(headers, "Submitted Date");
    const mobileIdx        = _colIndex(headers, "Mobile");

    const incomingClaimId = String(body[CLAIM_ID_COLUMN] || body["claim_id"] || "").trim();
    const incomingSrNo    = String(body[SR_NO_COLUMN]    || body["sr_no"]    || "").trim();
    const legacyLookup    = String(body["_legacy_sr_lookup"] || "").trim();

    Logger.log("incomingClaimId=" + incomingClaimId + " incomingSrNo=" + incomingSrNo + " legacyLookup=" + legacyLookup);

    const rowIndex = _findRowIndex(sheet, headers, incomingClaimId, incomingSrNo, legacyLookup);

    if (rowIndex !== -1) {
      // UPDATE existing row
      var existingRange  = sheet.getRange(rowIndex, 1, 1, headers.length);
      var existingValues = existingRange.getValues()[0];
      while (existingValues.length < headers.length) existingValues.push("");

      // Merge payload (skip blacklisted columns)
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i];
        if (SHEET_BLACKLIST.indexOf(h) !== -1) continue;
        if (h === "_legacy_sr_lookup") continue;
        if (body.hasOwnProperty(h)) existingValues[i] = body[h];
      }

      // Handle header aliases
      if (submittedDateIdx !== -1 && (body["Submitted Date"] || body["Date"])) {
        existingValues[submittedDateIdx] = body["Submitted Date"] || body["Date"];
      }
      if (mobileIdx !== -1 && (body["Mobile"] || body["Mobile Number"])) {
        existingValues[mobileIdx] = body["Mobile"] || body["Mobile Number"];
      }

      // Always write Claim ID to its dedicated column
      if (claimIdColIdx !== -1 && incomingClaimId) {
        existingValues[claimIdColIdx] = incomingClaimId;
      }

      // Write SR No: only real SR Nos (not CLM-XXXXX values)
      if (srNoColIdx !== -1) {
        if (incomingSrNo && !incomingSrNo.startsWith("CLM-")) {
          // Real SR No — write it
          existingValues[srNoColIdx] = incomingSrNo;
        } else if (!incomingSrNo) {
          // No SR No in this update — preserve existing value (do nothing)
        }
        // Clear CLM legacy from SR No if it was there
        if (existingValues[srNoColIdx] && String(existingValues[srNoColIdx]).startsWith("CLM-")) {
          existingValues[srNoColIdx] = "";
        }
      }

      sheet.getRange(rowIndex, 1, 1, existingValues.length).setValues([existingValues]);

      // Re-apply center alignment on update too
      const updatedRange = sheet.getRange(rowIndex, 1, 1, existingValues.length);
      updatedRange.setHorizontalAlignment("center");
      updatedRange.setVerticalAlignment("middle");

      if (submittedDateIdx !== -1) {
        sheet.getRange(rowIndex, submittedDateIdx + 1).setNumberFormat("@STRING@");
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "updated", row: rowIndex, claim_id: incomingClaimId, sr_no: incomingSrNo
      })).setMimeType(ContentService.MimeType.JSON);

    } else {
      // INSERT new row
      const rowData = new Array(headers.length).fill("");

      for (let i = 0; i < headers.length; i++) {
        const h = headers[i];
        if (SHEET_BLACKLIST.indexOf(h) !== -1) continue;
        if (h === "_legacy_sr_lookup") continue;
        if (body.hasOwnProperty(h)) rowData[i] = body[h];
      }

      // Handle header aliases
      if (submittedDateIdx !== -1 && (body["Submitted Date"] || body["Date"])) {
        rowData[submittedDateIdx] = body["Submitted Date"] || body["Date"];
      }
      if (mobileIdx !== -1 && (body["Mobile"] || body["Mobile Number"])) {
        rowData[mobileIdx] = body["Mobile"] || body["Mobile Number"];
      }

      // CLM -> Claim ID column (NOT SR No)
      if (claimIdColIdx !== -1 && incomingClaimId) rowData[claimIdColIdx] = incomingClaimId;

      // SR No: blank for new claims, real value only if explicitly provided
      if (srNoColIdx !== -1) {
        rowData[srNoColIdx] = (incomingSrNo && !incomingSrNo.startsWith("CLM-")) ? incomingSrNo : "";
      }

      sheet.appendRow(rowData);

      // Apply center alignment to the newly inserted row
      const newRowNum = sheet.getLastRow();
      const newRowRange = sheet.getRange(newRowNum, 1, 1, headers.length);
      newRowRange.setHorizontalAlignment("center");
      newRowRange.setVerticalAlignment("middle");

      if (submittedDateIdx !== -1) {
        sheet.getRange(newRowNum, submittedDateIdx + 1).setNumberFormat("@STRING@");
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "created", claim_id: incomingClaimId
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    Logger.log("ERROR: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: "error", message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------
// GET handler
// -----------------------------------------
function doGet(e) {
  try {
    const sheet   = _getSheet();
    const headers = _readHeaders(sheet);
    var   data    = sheet.getDataRange().getValues();
    if (data.length <= HEADER_ROW_INDEX) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    data = data.slice(HEADER_ROW_INDEX);
    const result = data.map(row => {
      const obj = {};
      for (let c = 0; c < headers.length; c++) {
        if (!headers[c]) continue;
        let val = row[c];
        if (val instanceof Date) val = Utilities.formatDate(val, "Asia/Kolkata", "yyyy-MM-dd");
        obj[headers[c]] = val;
      }
      return obj;
    });
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -----------------------------------------
// onEdit trigger for Webhook integration
// -----------------------------------------
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  
  const headers = _readHeaders(sheet);
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  
  if (editedRow <= HEADER_ROW_INDEX) return; // ignore header edits
  
  const headerName = headers[editedCol - 1];
  const lowerHeader = (headerName || "").toLowerCase();
  
  // Only trigger webhook if Remarks, ONSITEGO - STATUS, or SR No are edited
  if (lowerHeader === "remarks" || (lowerHeader.includes("onsitego") && lowerHeader.includes("status")) || lowerHeader === "sr no" || lowerHeader === "sr_no") {
    const claimIdIdx = _colIndex(headers, CLAIM_ID_COLUMN);
    if (claimIdIdx === -1) return;
    
    // Read the whole row to send to webhook
    const rowValues = sheet.getRange(editedRow, 1, 1, headers.length).getValues()[0];
    const claimId = rowValues[claimIdIdx];
    
    if (!claimId) return; // No claim ID, cannot sync back
    
    const payload = {};
    for (let i = 0; i < headers.length; i++) {
      if (headers[i]) {
        let val = rowValues[i];
        if (val instanceof Date) val = Utilities.formatDate(val, "Asia/Kolkata", "yyyy-MM-dd");
        payload[headers[i]] = val;
      }
    }
    
    // IMPORTANT: Replace this with your actual portal domain (e.g. https://my-portal.onrender.com)
    const PORTAL_URL = "YOUR_PORTAL_URL_HERE"; 
    
    if (PORTAL_URL === "YOUR_PORTAL_URL_HERE") {
      Logger.log("Please update PORTAL_URL in the Apps Script to sync back to the portal.");
      return;
    }
    
    const webhookUrl = PORTAL_URL + "/api/webhook/sheet-update";
    
    try {
      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      UrlFetchApp.fetch(webhookUrl, options);
    } catch (err) {
      Logger.log("Failed to hit webhook: " + err);
    }
  }
}
