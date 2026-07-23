var DRIVE_FOLDER_ID = "1r0i_-BcGPbSguJLAXMD9MtaojaxRgUOM";

function uploadFileToDrive(base64Data, fileName, mimeType) {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType, fileName);
    var file = folder.createFile(blob);
    
    // Set file sharing access so that users with the link can view/download
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      fileName: file.getName(),
      fileSize: file.getSize()
    };
  } catch (err) {
    throw new Error("File upload failed: " + err.toString());
  }
}

/**
 * Converts a compiled DOCX file (sent as base64) directly into a native PDF
 * using Google Drive's native document layout & export engine.
 * Converts .docx -> Google Doc -> PDF to preserve 100% Word formatting.
 */
function convertDocxToPdf(docxBase64, fileName) {
  try {
    var name = fileName || "Consolidated_KRA";
    var bytes = Utilities.base64Decode(docxBase64);
    
    // Upload .docx to Google Drive with convert=true to import as Google Doc
    var url = "https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart&convert=true";
    var boundary = "---" + Math.random().toString(16).substring(2);
    
    var metadata = JSON.stringify({
      title: name,
      mimeType: "application/vnd.google-apps.document",
      parents: [{ id: DRIVE_FOLDER_ID }]
    });
    
    var requestBody = 
      "--" + boundary + "\r\n" +
      "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
      metadata + "\r\n" +
      "--" + boundary + "\r\n" +
      "Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n";
    
    var postData = Utilities.newBlob(requestBody).getBytes();
    postData = postData.concat(bytes);
    
    var closingBoundary = Utilities.newBlob("\r\n--" + boundary + "--").getBytes();
    postData = postData.concat(closingBoundary);
    
    var options = {
      method: "post",
      contentType: "multipart/related; boundary=" + boundary,
      payload: postData,
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    if (!result.id) {
      throw new Error("Drive conversion error: " + response.getContentText());
    }
    
    // Fetch converted Google Doc as PDF
    var googleDocFile = DriveApp.getFileById(result.id);
    var pdfBlob = googleDocFile.getAs("application/pdf");
    pdfBlob.setName(name + ".pdf");
    
    // Encode to base64
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    // Trash the temporary Google Doc file
    googleDocFile.setTrashed(true);
    
    return {
      pdfBase64: pdfBase64,
      fileName: name + ".pdf"
    };
  } catch (err) {
    throw new Error("DOCX to PDF conversion failed: " + err.toString());
  }
}
