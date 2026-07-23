/**
 * Admin Service for Quarter Management, ITA Certificates, and Audit Logs
 */

/**
 * Locks a specific quarter for a financial year (Admin only).
 * @param {string} token 
 * @param {string} fy 
 * @param {string} quarter 
 * @returns {object} {success: true}
 */
function lockQuarter(token, fy, quarter) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var qsSheet = ss.getSheetByName('Quarter_Settings');
  var data = qsSheet.getDataRange().getValues();
  
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === fy && data[i][1] === quarter) {
      qsSheet.getRange(i + 1, 3).setValue(true); // IsLocked
      qsSheet.getRange(i + 1, 4).setValue(sessionUser.username); // LockedBy
      qsSheet.getRange(i + 1, 5).setValue(new Date()); // LockedDate
      found = true;
      break;
    }
  }
  
  if (!found) {
    qsSheet.appendRow([fy, quarter, true, sessionUser.username, new Date(), ""]);
  }
  
  logAuditEntry("LOCK_QUARTER", "ALL", quarter, fy, sessionUser.username, sessionUser.role, "UNLOCKED", "LOCKED", "Quarter locked");
  
  return { success: true };
}

/**
 * Unlocks a specific quarter for a financial year (Admin only).
 * @param {string} token 
 * @param {string} fy 
 * @param {string} quarter 
 * @returns {object} {success: true}
 */
function unlockQuarter(token, fy, quarter) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var qsSheet = ss.getSheetByName('Quarter_Settings');
  var data = qsSheet.getDataRange().getValues();
  
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === fy && data[i][1] === quarter) {
      qsSheet.getRange(i + 1, 3).setValue(false); // IsLocked
      qsSheet.getRange(i + 1, 4).setValue(sessionUser.username); // LockedBy (or could be cleared)
      qsSheet.getRange(i + 1, 5).setValue(new Date()); // LockedDate (or unlocked date)
      found = true;
      break;
    }
  }
  
  if (!found) {
    qsSheet.appendRow([fy, quarter, false, sessionUser.username, new Date(), ""]);
  }
  
  logAuditEntry("UNLOCK_QUARTER", "ALL", quarter, fy, sessionUser.username, sessionUser.role, "LOCKED", "UNLOCKED", "Quarter unlocked");
  
  return { success: true };
}

/**
 * Gets all quarter settings.
 * @returns {Array} List of quarter settings
 */
function getQuarterSettings() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var qsSheet = ss.getSheetByName('Quarter_Settings');
  if (!qsSheet) return [];
  
  var data = qsSheet.getDataRange().getValues();
  var settings = [];
  
  for (var i = 1; i < data.length; i++) {
    settings.push({
      fy: data[i][0],
      quarter: data[i][1],
      isLocked: data[i][2],
      lockedBy: data[i][3],
      lockedDate: data[i][4],
      deadline: data[i][5]
    });
  }
  
  return settings;
}

/**
 * Checks if a specific quarter is locked.
 * @param {string} fy 
 * @param {string} quarter 
 * @returns {boolean} True if locked, false otherwise
 */
function isQuarterLocked(fy, quarter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var qsSheet = ss.getSheetByName('Quarter_Settings');
  if (!qsSheet) return false;
  
  var data = qsSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === fy && data[i][1] === quarter) {
      return data[i][2] === true;
    }
  }
  return false;
}

/**
 * Saves an ITA certificate (ITA/Admin only).
 * @param {string} token 
 * @param {object} certData {fy, quarter, observations, certifiedBy, certifiedDate}
 * @returns {object} {success: true}
 */
function saveITACertificate(token, certData) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'ITA_OFFICER' && sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. ITA Officer or Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var itaSheet = ss.getSheetByName('ITA_Certificates');
  var data = itaSheet.getDataRange().getValues();
  
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === certData.fy && data[i][1] === certData.quarter) {
      itaSheet.getRange(i + 1, 3).setValue(certData.certifiedBy);
      itaSheet.getRange(i + 1, 4).setValue(certData.certifiedDate || new Date());
      itaSheet.getRange(i + 1, 5).setValue(certData.observations);
      itaSheet.getRange(i + 1, 6).setValue("CERTIFIED");
      found = true;
      break;
    }
  }
  
  if (!found) {
    itaSheet.appendRow([
      certData.fy, 
      certData.quarter, 
      certData.certifiedBy, 
      certData.certifiedDate || new Date(), 
      certData.observations, 
      "CERTIFIED"
    ]);
  }
  
  logAuditEntry("SAVE_ITA_CERTIFICATE", "ITA", certData.quarter, certData.fy, sessionUser.username, sessionUser.role, "", "CERTIFIED", "ITA Certificate saved");
  
  return { success: true };
}

/**
 * Gets the ITA certificate for a specific quarter.
 * @param {string} fy 
 * @param {string} quarter 
 * @returns {object|null} Certificate data or null
 */
function getITACertificate(fy, quarter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var itaSheet = ss.getSheetByName('ITA_Certificates');
  if (!itaSheet) return null;
  
  var data = itaSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === fy && data[i][1] === quarter) {
      return {
        fy: data[i][0],
        quarter: data[i][1],
        certifiedBy: data[i][2],
        certifiedDate: data[i][3],
        observations: data[i][4],
        status: data[i][5]
      };
    }
  }
  return null;
}

/**
 * Gets audit log entries (Admin only).
 * @param {string} token 
 * @param {object} filters Optional filters {startDate, endDate, section, action}
 * @returns {Array} List of audit logs
 */
function getAuditLog(token, filters) {
  var sessionUser = validateSession(token);
  if (!sessionUser || (sessionUser.role !== 'TM_ADMIN' && sessionUser.role !== 'ADMIN' && sessionUser.role !== 'DEVELOPER')) {
    throw new Error("Unauthorized. Admin access required.");
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var logSheet = ss.getSheetByName('Audit_Log');
  if (!logSheet) return [];
  
  var data = logSheet.getDataRange().getValues();
  var logs = [];
  
  for (var i = 1; i < data.length; i++) {
    var timestamp = data[i][0];
    var action = data[i][1];
    var section = data[i][2];
    
    // Apply filters if present
    if (filters) {
      if (filters.startDate && new Date(timestamp) < new Date(filters.startDate)) continue;
      if (filters.endDate && new Date(timestamp) > new Date(filters.endDate)) continue;
      if (filters.section && section !== filters.section) continue;
      if (filters.action && action !== filters.action) continue;
    }
    
    logs.push({
      timestamp: timestamp,
      action: action,
      section: section,
      quarter: data[i][3],
      fy: data[i][4],
      performedBy: data[i][5],
      role: data[i][6],
      previousStatus: data[i][7],
      newStatus: data[i][8],
      comments: data[i][9]
    });
  }
  
  return logs;
}

/**
 * Internal helper to write an entry to the Audit Log.
 */
function logAuditEntry(action, section, quarter, fy, performedBy, role, prevStatus, newStatus, comments) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var logSheet = ss.getSheetByName('Audit_Log');
  if (logSheet) {
    logSheet.appendRow([
      new Date(),
      action,
      section,
      quarter,
      fy,
      performedBy,
      role,
      prevStatus,
      newStatus,
      comments || ""
    ]);
  }
}
