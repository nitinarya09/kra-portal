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
 */
function convertDocxToPdf(docxBase64, fileName) {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var name = fileName || "Consolidated_KRA";
    var bytes = Utilities.base64Decode(docxBase64);
    var docxBlob = Utilities.newBlob(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", name + ".docx");
    
    // Create temporary DOCX file in Drive
    var tempFile = folder.createFile(docxBlob);
    
    // Convert to PDF natively via Drive layout engine
    var pdfBlob = tempFile.getAs("application/pdf");
    pdfBlob.setName(name + ".pdf");
    
    // Encode to base64
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    // Trash the temporary file so Drive remains clean
    tempFile.setTrashed(true);
    
    return {
      pdfBase64: pdfBase64,
      fileName: name + ".pdf"
    };
  } catch (err) {
    throw new Error("DOCX to PDF conversion failed: " + err.toString());
  }
}
