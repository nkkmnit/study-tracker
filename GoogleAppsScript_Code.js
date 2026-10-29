const SHEET_NAME = "Sheet1";

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const p = (e && e.parameter) ? e.parameter : {};
    
    // Action: Register via GET (Bulletproof CORS fallback)
    if (p.action === "register") {
      const deviceId = p.deviceId;
      const name = p.name;
      const phone = p.phone;
      const amount = p.amount || 199;
      const date = p.date || new Date().toLocaleDateString('hi-IN');
      
      const rows = sheet.getDataRange().getValues();
      let exists = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(deviceId).trim()) {
          exists = true;
          break;
        }
      }
      if (!exists) {
        sheet.appendRow([deviceId, name, phone, amount, date, "pending"]);
      }
      return outputJSON({ status: "success", message: "registered" });
    }
    
    // Action: Update Status via GET
    if (p.action === "updateStatus") {
      const deviceId = p.deviceId;
      const newStatus = p.newStatus;
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(deviceId).trim()) {
          sheet.getRange(i + 1, 6).setValue(newStatus);
          break;
        }
      }
      return outputJSON({ status: "updated", deviceId: deviceId, newStatus: newStatus });
    }
    
    // Default GET: Fetch all user records
    const rows = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) {
        result.push({
          deviceId: String(rows[i][0]),
          name: String(rows[i][1] || 'Anonymous'),
          phone: String(rows[i][2] || 'N/A'),
          amount: rows[i][3] || 199,
          date: String(rows[i][4] || ''),
          status: String(rows[i][5] || 'pending').toLowerCase().trim()
        });
      }
    }
    return outputJSON(result);
  } catch (err) {
    return outputJSON({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    let data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    
    if (data.action === "register") {
      const rows = sheet.getDataRange().getValues();
      let exists = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(data.deviceId).trim()) {
          exists = true;
          break;
        }
      }
      if (!exists) {
        sheet.appendRow([data.deviceId, data.name, data.phone, data.amount, data.date, "pending"]);
      }
      return outputJSON({ status: "success" });
    }
    
    if (data.action === "updateStatus") {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(data.deviceId).trim()) {
          sheet.getRange(i + 1, 6).setValue(data.newStatus);
          break;
        }
      }
      return outputJSON({ status: "updated" });
    }
    return outputJSON({ status: "no_action" });
  } catch (err) {
    return outputJSON({ status: "error", message: err.toString() });
  }
}

function outputJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
