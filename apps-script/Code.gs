var DRIVE_FOLDER_ID = "1r0i_-BcGPbSguJLAXMD9MtaojaxRgUOM";
var SPREADSHEET_ID = "12YfeR994vYcuQTdvA1y-8l-dXJzX64AlP9GG8EtSEjM";

function doPost(e) {
  try {
    var request = JSON.parse(e.postData.contents);
    var action = request.action;
    
    var response;
    if (action === "submitForm") {
      if (request.payload.token && !validateSession(request.payload.token)) throw new Error("Unauthorized session.");
      if (isQuarterLocked(request.payload.metadata.fy, request.payload.metadata.quarter)) throw new Error("Quarter is locked for submissions.");
      response = submitFormData(request.payload);
    } else if (action === "uploadExcel") {
      if (request.payload.token && !validateSession(request.payload.token)) throw new Error("Unauthorized session.");
      response = parseAndSaveExcel(request.payload);
    } else if (action === "uploadFile") {
      response = uploadFileToDrive(request.payload.base64Data, request.payload.fileName, request.payload.mimeType);
    } else if (action === "getPreviousClosing") {
      response = getPreviousClosing(request.payload);
    } else if (action === "getSubmissions") {
      response = getSubmissionLog();
    } else if (action === "login") {
      response = authenticateUser(request.payload.username, request.payload.passwordHash);
    } else if (action === "logout") {
      response = logoutUser(request.payload.token);
    } else if (action === "changePassword") {
      response = changePassword(request.payload.token, request.payload.oldHash, request.payload.newHash);
    } else if (action === "createUser") {
      response = createUser(request.payload.token, request.payload.userData);
    } else if (action === "updateUser") {
      response = updateUser(request.payload.token, request.payload.username, request.payload.updates);
    } else if (action === "deleteUser") {
      response = deleteUser(request.payload.token, request.payload.username);
    } else if (action === "resetPassword") {
      response = resetPassword(request.payload.token, request.payload.targetUsername, request.payload.newHash);
    } else if (action === "listUsers") {
      response = listUsers(request.payload.token);
    } else if (action === "seedDefaultAccounts") {
      response = seedDefaultAccounts(request.payload.token);
    } else if (action === "lockQuarter") {
      response = lockQuarter(request.payload.token, request.payload.fy, request.payload.quarter);
    } else if (action === "unlockQuarter") {
      response = unlockQuarter(request.payload.token, request.payload.fy, request.payload.quarter);
    } else if (action === "getQuarterSettings") {
      response = getQuarterSettings();
    } else if (action === "saveITACertificate") {
      response = saveITACertificate(request.payload.token, request.payload.certData);
    } else if (action === "getITACertificate") {
      response = getITACertificate(request.payload.fy, request.payload.quarter);
    } else if (action === "getAuditLog") {
      response = getAuditLog(request.payload.token, request.payload.filters);
    } else if (action === "approveSubmission") {
      response = approveSubmission(request.payload.token, request.payload.section, request.payload.fy, request.payload.quarter);
    } else if (action === "returnSubmission") {
      response = returnSubmission(request.payload.token, request.payload.section, request.payload.fy, request.payload.quarter, request.payload.comments);
    } else if (action === "getAllKRAData") {
      response = getAllKRAData(request.payload);
    } else if (action === "convertDocxToPdf") {
      response = convertDocxToPdf(request.payload.docxBase64, request.payload.fileName);
    } else {
      throw new Error("Invalid action code");
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", data: response }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: err.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", message: "API is active" }))
                       .setMimeType(ContentService.MimeType.JSON);
}
