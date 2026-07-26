function submitFormData(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var section = payload.metadata.section;
  var fy = payload.metadata.fy;
  var quarter = payload.metadata.quarter;
  
  if (typeof isQuarterLocked === 'function' && isQuarterLocked(fy, quarter)) {
    throw new Error("Quarter is locked for submissions.");
  }
  
  // 1. Clear existing draft/submission for this Section+FY+Quarter in all affected sheets
  var domainSheets = Object.keys(payload.data);
  for (var i = 0; i < domainSheets.length; i++) {
    var sheet = ss.getSheetByName(domainSheets[i]);
    if (sheet) {
      removeExistingRecords(sheet, section, fy, quarter);
      
      // Append new data rows
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var rows = payload.data[domainSheets[i]];
      for (var r = 0; r < rows.length; r++) {
        var rowData = headers.map(function(h) {
          if (h === "Section") return section;
          if (h === "FY") return fy;
          if (h === "Quarter") return quarter;
          return rows[r][h] !== undefined ? rows[r][h] : "";
        });
        sheet.appendRow(rowData);
      }
    }
  }
  
  // 2. Update Master Submissions Log
  var subSheet = ss.getSheetByName("Submissions");
  removeExistingRecords(subSheet, section, fy, quarter);
  subSheet.appendRow([
    section, fy, quarter, "SUBMITTED",
    payload.metadata.submittedBy, payload.metadata.designation,
    new Date(), 1, payload.metadata.remarks || "", "", "", ""
  ]);
  
  // 3. Update Attachments File Registry if any
  if (payload.attachments && payload.attachments.length > 0) {
    var attSheet = ss.getSheetByName("Attachments");
    removeExistingRecords(attSheet, section, fy, quarter);
    for (var a = 0; a < payload.attachments.length; a++) {
      var att = payload.attachments[a];
      attSheet.appendRow([
        section, fy, quarter, att.fileName, att.driveUrl, att.fileSize, new Date(), att.annexure || ""
      ]);
    }
  }
  
  return { section: section, status: "Submitted successfully" };
}

function removeExistingRecords(sheet, section, fy, quarter) {
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === section && data[i][1] === fy && data[i][2] === quarter) {
      sheet.deleteRow(i + 1);
    }
  }
}

function getPreviousClosing(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var section = payload.section;
  var prevQuarter = payload.prevQuarter; // e.g. "Q4"
  var prevFy = payload.prevFy;           // e.g. "2025-26"
  
  var result = {};
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    if (sheetName === "Submissions" || sheetName === "Attachments") continue;
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue;
    var headers = data[0];
    
    // Find matching rows for this section + previous quarter + previous FY
    var sectionRows = [];
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] === section && data[r][1] === prevFy && data[r][2] === prevQuarter) {
        var rowObj = {};
        for (var c = 0; c < headers.length; c++) {
          rowObj[headers[c]] = data[r][c];
        }
        sectionRows.push(rowObj);
      }
    }
    if (sectionRows.length > 0) {
      result[sheetName] = sectionRows;
    }
  }
  return result;
}

function getSubmissionLog() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Submissions");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var logs = [];
  
  for (var r = 1; r < data.length; r++) {
    var logObj = {};
    for (var c = 0; c < headers.length; c++) {
      var val = data[r][c];
      // Format Date values cleanly
      if (val instanceof Date) {
        logObj[headers[c]] = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
      } else {
        logObj[headers[c]] = val;
      }
    }
    logs.push(logObj);
  }
  return logs;
}

function approveSubmission(token, section, fy, quarter) {
  var sessionUser = validateSession(token);
  if (!sessionUser || sessionUser.role === 'OPERATOR') {
    throw new Error("Unauthorized. Reviewer access or above required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var subSheet = ss.getSheetByName("Submissions");
  var data = subSheet.getDataRange().getValues();
  
  var newStatus = "APPROVED_L1"; // Default for REVIEWER
  if (sessionUser.role === 'SECTION_HEAD' || sessionUser.role === 'TM_ADMIN' || sessionUser.role === 'ADMIN' || sessionUser.role === 'DEVELOPER' || sessionUser.role === 'ITA_OFFICER') {
    newStatus = "APPROVED_L2";
  }
  
  var found = false;
  var prevStatus = "";
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === section && data[i][1] === fy && data[i][2] === quarter) {
      prevStatus = data[i][3];
      subSheet.getRange(i + 1, 4).setValue(newStatus); // Status
      subSheet.getRange(i + 1, 10).setValue(sessionUser.username); // ReviewedBy
      subSheet.getRange(i + 1, 11).setValue(new Date()); // ReviewDate
      subSheet.getRange(i + 1, 12).setValue(""); // Clear return comments
      found = true;
      break;
    }
  }
  
  if (!found) throw new Error("Submission not found.");
  
  if (typeof logAuditEntry === 'function') {
    logAuditEntry("APPROVE_SUBMISSION", section, quarter, fy, sessionUser.username, sessionUser.role, prevStatus, newStatus, "");
  }
  
  return { success: true, status: newStatus };
}

function returnSubmission(token, section, fy, quarter, comments) {
  var sessionUser = validateSession(token);
  if (!sessionUser || sessionUser.role === 'OPERATOR') {
    throw new Error("Unauthorized. Reviewer access or above required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var subSheet = ss.getSheetByName("Submissions");
  var data = subSheet.getDataRange().getValues();
  
  var found = false;
  var prevStatus = "";
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === section && data[i][1] === fy && data[i][2] === quarter) {
      prevStatus = data[i][3];
      subSheet.getRange(i + 1, 4).setValue("RETURNED"); // Status
      subSheet.getRange(i + 1, 10).setValue(sessionUser.username); // ReviewedBy
      subSheet.getRange(i + 1, 11).setValue(new Date()); // ReviewDate
      subSheet.getRange(i + 1, 12).setValue(comments); // ReturnComments
      found = true;
      break;
    }
  }
  
  if (!found) throw new Error("Submission not found.");
  
  if (typeof logAuditEntry === 'function') {
    logAuditEntry("RETURN_SUBMISSION", section, quarter, fy, sessionUser.username, sessionUser.role, prevStatus, "RETURNED", comments);
  }
  
  return { success: true, status: "RETURNED" };
}

function getAllKRAData(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var fy = payload.fy;
  var quarter = payload.quarter || "Q1";
  var qName = quarter.split(" ")[0]; // e.g. "Q1"
  
  var result = {};
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    if (sheetName === "Submissions" || sheetName === "Attachments" || sheetName === "Users" || sheetName === "Sessions" || sheetName === "Audit_Log") continue;
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue;
    var headers = data[0];
    
    var matchingRows = [];
    for (var r = 1; r < data.length; r++) {
      var rowFy = String(data[r][1] || '').trim();
      var rowQtr = String(data[r][2] || '').trim();
      if ((!rowFy || rowFy === fy) && (!rowQtr || rowQtr === qName || rowQtr.indexOf(qName) === 0)) {
        var rowObj = {};
        for (var c = 0; c < headers.length; c++) {
          rowObj[headers[c]] = data[r][c];
        }
        matchingRows.push(rowObj);
      }
    }
    if (matchingRows.length > 0) {
      result[sheetName] = matchingRows;
    }
  }
  return result;
}
