// --------------------------------------------------------------------------------
// OSG myG PORTAL - BACKEND SCRIPT
// --------------------------------------------------------------------------------
// INSTRUCTIONS:
// 1. Paste this code into your Google Apps Script editor.
// 2. Save the project.
// 3. Click "Deploy" -> "New Deployment".
// 4. Select type: "Web App".
// 5. Description: "v2 Upsert logic".
// 6. Execute as: "Me" (your email).
// 7. Who has access: "Anyone" (IMPORTANT: Do not select "Anyone with Google Account" or "Only me").
// 8. Click "Deploy" and copy the NEW Web App URL.
// 9. Update the 'WEB_APP_URL' in your Python 'app.py' file if the URL changed.
// --------------------------------------------------------------------------------

const SHEET_NAME = "myG_OSG_Portal_Data";
const HEADER_ROW_INDEX = 1;

function _getSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
        // Auto-create if missing
        sheet = ss.insertSheet(SHEET_NAME);
        // Add default headers if new
        sheet.appendRow([
            "Claim ID", "Date", "Customer Name", "Mobile Number", "Address",
            "Product", "Invoice Number", "Serial Number", "Model", "OSID", "Issue", "Branch",
            "Follow Up - Dates", "Follow Up - Notes", "Claim Settled Date", "Remarks",
            "Status",
            // Replacement Workflow Columns (O-T)
            "Replacement: Confirmation Pending",
            "Replacement: OSG Approval",
            "Replacement: Mail to Store",
            "Replacement: Invoice Generated",
            "Replacement: Invoice Sent to OSG",
            "Replacement: Settled with Accounts",
            // Complete Flag (Column U)
            "Complete",
            // Other fields
            "Settled Time (TAT)", "Assigned Staff", "Last Updated Timestamp"
        ]);
    }
    return sheet;
}

function _readHeaders(sheet) {
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const raw = sheet.getRange(HEADER_ROW_INDEX, 1, 1, lastCol).getValues()[0] || [];
    return raw.map(h => (h === null || h === undefined) ? "" : String(h).trim());
}

function _findHeaderIndex(headers, nameOptions) {
    if (!Array.isArray(nameOptions)) nameOptions = [nameOptions];
    const opts = nameOptions.map(x => String(x).trim().toLowerCase());
    for (let i = 0; i < headers.length; i++) {
        const h = (headers[i] || "").toString().trim().toLowerCase();
        if (opts.indexOf(h) !== -1) return i;
    }
    // partial match fallback
    for (let i = 0; i < headers.length; i++) {
        const h = (headers[i] || "").toString().trim().toLowerCase();
        for (const opt of opts) if (opt && h.indexOf(opt) !== -1) return i;
    }
    return -1;
}

function _generateNextClaimId(sheet, claimIdColIndex) {
    // If we can't find column, just time-based
    if (claimIdColIndex < 0) return "CLM-" + new Date().getTime();

    const startRow = HEADER_ROW_INDEX + 1;
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return "CLM-0001";

    const vals = sheet.getRange(startRow, claimIdColIndex + 1, lastRow - HEADER_ROW_INDEX + 1, 1).getValues().flat();
    let maxN = 0;
    vals.forEach(v => {
        const s = String(v || "");
        const parts = s.split("-");
        if (parts.length > 1) {
            const n = parseInt(parts[1], 10);
            if (!isNaN(n) && n > maxN) maxN = n;
        }
    });
    return "CLM-" + String(maxN + 1).padStart(4, "0");
}

function _currentIstIso() {
    return Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
}

function doPost(e) {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000); // Wait up to 10s for other concurrent requests

    try {
        var body = {};
        if (e.postData && e.postData.contents) {
            body = JSON.parse(e.postData.contents);
        } else if (e.parameter) {
            body = e.parameter;
        }

        var sheet = _getSheet();
        var headers = _readHeaders(sheet);

        // Map critical columns
        var colMap = {
            "Claim ID": _findHeaderIndex(headers, ["Claim ID", "claimid", "id"]),
            "Date": _findHeaderIndex(headers, ["Date", "Submitted Date"]),
            "Customer Name": _findHeaderIndex(headers, ["Customer Name", "customer"]),
            "Mobile Number": _findHeaderIndex(headers, ["Mobile Number", "mobile no", "mobile"]),
            "Status": _findHeaderIndex(headers, ["Status"]),
            "Last Updated": _findHeaderIndex(headers, ["Last Updated Timestamp", "last updated"])
        };

        // Determine Claim ID
        var incomingId = body["Claim ID"] || body["claim_id"];

        // If incomingId is missing, generate one (New Submission without ID from Python?)
        // Python usually sends "CLM-..."
        if (!incomingId && colMap["Claim ID"] !== -1) {
            incomingId = _generateNextClaimId(sheet, colMap["Claim ID"]);
        }

        // Construct the Row Data Array
        var rowData = new Array(headers.length).fill("");

        // Helper to fill row data
        function fillRow(r) {
            for (var i = 0; i < headers.length; i++) {
                var headerName = headers[i];
                // Check exact match in payload
                if (body.hasOwnProperty(headerName)) {
                    r[i] = body[headerName];
                } else {
                    // Check normalized keys match (e.g. "claim_id" -> "Claim ID")
                    // (Skipping complex normalization for now, relying on Python to send exact headers is safer
                    //  OR simple keys if Python sends simple keys. Python currently sends "Claim ID" etc.)
                }
            }
            // Explicit overrides if keys differ slightly in body
            if (colMap["Claim ID"] !== -1) r[colMap["Claim ID"]] = incomingId;
            if (colMap["Last Updated"] !== -1) r[colMap["Last Updated"]] = _currentIstIso();
            return r;
        }

        rowData = fillRow(rowData);

        // UPSERT LOGIC: Check if ID exists
        var claimIdIdx = colMap["Claim ID"];
        var rowIndexToUpdate = -1;

        if (claimIdIdx !== -1 && incomingId) {
            var lastRow = sheet.getLastRow();
            if (lastRow > HEADER_ROW_INDEX) {
                var ids = sheet.getRange(HEADER_ROW_INDEX + 1, claimIdIdx + 1, lastRow - HEADER_ROW_INDEX, 1).getValues().flat();
                for (var i = 0; i < ids.length; i++) {
                    if (String(ids[i]) === String(incomingId)) {
                        rowIndexToUpdate = HEADER_ROW_INDEX + 1 + i;
                        break;
                    }
                }
            }
        }

        if (rowIndexToUpdate !== -1) {
            // Update existing row
            // FETCH EXISTING DATA FIRST to avoid overwriting with blanks
            var existingValues = sheet.getRange(rowIndexToUpdate, 1, 1, headers.length).getValues()[0];

            // Merge existing with new
            for (var i = 0; i < headers.length; i++) {
                var headerName = headers[i];
                if (body.hasOwnProperty(headerName)) {
                    existingValues[i] = body[headerName];
                }
            }

            // Still update forced fields
            if (colMap["Last Updated"] !== -1) existingValues[colMap["Last Updated"]] = _currentIstIso();

            sheet.getRange(rowIndexToUpdate, 1, 1, existingValues.length).setValues([existingValues]);

            // Format Date column as plain text to prevent auto-formatting
            var dateColIdx = colMap["Date"];
            if (dateColIdx !== -1) {
                sheet.getRange(rowIndexToUpdate, dateColIdx + 1).setNumberFormat('@STRING@');
            }

            return ContentService.createTextOutput(JSON.stringify({
                "status": "updated",
                "row": rowIndexToUpdate,
                "claim_id": incomingId
            })).setMimeType(ContentService.MimeType.JSON);

        } else {
            // Append new row - Use the constructed rowData which is based on payload + blanks
            // (Since it's new, blanks are fine for missing fields)
            sheet.appendRow(rowData);

            // Format Date column as plain text to prevent auto-formatting
            var dateColIdx = colMap["Date"];
            if (dateColIdx !== -1) {
                var newRowIndex = sheet.getLastRow();
                sheet.getRange(newRowIndex, dateColIdx + 1).setNumberFormat('@STRING@');
            }

            return ContentService.createTextOutput(JSON.stringify({
                "status": "created",
                "claim_id": incomingId
            })).setMimeType(ContentService.MimeType.JSON);
        }

    } catch (e) {
        return ContentService.createTextOutput(JSON.stringify({
            "status": "error",
            "message": e.toString()
        })).setMimeType(ContentService.MimeType.JSON);

    } finally {
        lock.releaseLock();
    }
}

function doGet(e) {
    try {
        var sheet = _getSheet();
        var headers = _readHeaders(sheet);
        var data = sheet.getDataRange().getValues();

        // Remove headers
        if (data.length > HEADER_ROW_INDEX) {
            data = data.slice(HEADER_ROW_INDEX);
        } else {
            return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
        }

        // Find Date column index
        var dateColIndex = _findHeaderIndex(headers, ["Date", "Submitted Date"]);

        var result = [];
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var obj = {};
            for (var c = 0; c < headers.length; c++) {
                var header = headers[c];
                // Normalize header to key if needed, or keep original
                // Using original headers ensures consistency with the Sheet
                if (header) {
                    var value = row[c];

                    // Format Date columns as YYYY-MM-DD strings
                    if (c === dateColIndex && value instanceof Date) {
                        value = Utilities.formatDate(value, "Asia/Kolkata", "yyyy-MM-dd");
                    } else if (value instanceof Date) {
                        // Format any other Date objects
                        value = Utilities.formatDate(value, "Asia/Kolkata", "yyyy-MM-dd");
                    }

                    obj[header] = value;
                }
            }
            result.push(obj);
        }

        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    } catch (e) {
        return ContentService.createTextOutput(JSON.stringify({ "error": e.toString() })).setMimeType(ContentService.MimeType.JSON);
    }
}
