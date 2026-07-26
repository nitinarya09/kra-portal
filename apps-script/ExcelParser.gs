function parseAndSaveExcel(payload) {
  try {
    var bytes = Utilities.base64Decode(payload.base64File);
    var blob = Utilities.newBlob(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", payload.fileName);
    
    // 1. Upload temporary file to Drive and convert it to Google Sheets (Drive API v3)
    var resource = {
      name: "temp_convert_" + new Date().getTime(),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [DRIVE_FOLDER_ID]
    };
    
    // Convert blob to Google Sheets format
    var tempFile = Drive.Files.create(resource, blob);
    var tempFileId = tempFile.id;
    
    var tempSs = SpreadsheetApp.openById(tempFileId);
    var namedRanges = tempSs.getNamedRanges();
    var extractedData = {};
    
    // 2. Extract values from named ranges
    for (var i = 0; i < namedRanges.length; i++) {
      var range = namedRanges[i].getRange();
      var name = namedRanges[i].getName();
      var value = range.getValue();
      extractedData[name] = value;
    }
    
    // 3. Clean up the temporary spreadsheet
    Drive.Files.remove(tempFileId);
    
    // 4. Map the extracted named ranges to our worksheets
    var mappedData = mapExtractedNamedRanges(payload.metadata.section, extractedData);
    
    // 5. Save the mapped data
    var savePayload = {
      metadata: payload.metadata,
      data: mappedData,
      attachments: payload.attachments || []
    };
    
    return submitFormData(savePayload);
  } catch (err) {
    throw new Error("Excel parsing failed: " + err.toString());
  }
}

// Translate flat named ranges from the Excel template into worksheet row objects
function mapExtractedNamedRanges(section, data) {
  var mapped = {};
  
  // Example for Book-1 Excel structure:
  if (section === "BOOK1") {
    mapped["Accounts_MCA"] = [{
      "Month1": data["MCA_Month1"] || "",
      "Date1": data["MCA_Date1"] || "",
      "Delay1": data["MCA_Delay1"] || 0,
      "Month2": data["MCA_Month2"] || "",
      "Date2": data["MCA_Date2"] || "",
      "Delay2": data["MCA_Delay2"] || 0,
      "Month3": data["MCA_Month3"] || "",
      "Date3": data["MCA_Date3"] || "",
      "Delay3": data["MCA_Delay3"] || 0,
      "ExclPct1": data["MCA_ExclPct1"] || 0,
      "ExclPct2": data["MCA_ExclPct2"] || 0,
      "ExclPct3": data["MCA_ExclPct3"] || 0
    }];
    
    mapped["MKI_Dates"] = [
      { "Month": data["MKI_Month1"] || "", "UploadDate": data["MKI_UploadDate1"] || "" },
      { "Month": data["MKI_Month2"] || "", "UploadDate": data["MKI_UploadDate2"] || "" },
      { "Month": data["MKI_Month3"] || "", "UploadDate": data["MKI_UploadDate3"] || "" }
    ];
    
    // (Named ranges for Suspense details Table 8)
    mapped["Suspense_Remittance"] = [{
      "Head": "8793-Inter State Suspense",
      "SubHead": "Dr",
      "OB": data["ISS_Dr_OB"] || 0,
      "Addition": data["ISS_Dr_Add"] || 0,
      "Total": data["ISS_Dr_Total"] || 0,
      "Clearance": data["ISS_Dr_Clearance"] || 0,
      "ClearPct": data["ISS_Dr_ClearPct"] || 0,
      "OldCleared": data["ISS_Dr_OldCleared"] || 0,
      "CB": data["ISS_Dr_CB"] || 0
    }];
  }
  
  // (GEMINI: Add named range mapper rules for other sections here)
  
  return mapped;
}
