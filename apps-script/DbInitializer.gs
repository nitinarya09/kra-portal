function initializeDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Define worksheets and columns
  var sheets = {
    "Submissions": ["Section", "FY", "Quarter", "Status", "SubmittedBy", "Designation", "Timestamp", "Version", "Remarks", "ReviewedBy", "ReviewDate", "ReturnComments"],
    "Users": ["Username", "PasswordHash", "DisplayName", "Role", "AssignedSections", "Designation", "IsActive", "CreatedDate", "CreatedBy"],
    "Sessions": ["Token", "Username", "Role", "Sections", "LoginTime", "ExpiryTime"],
    "Login_Attempts": ["Timestamp", "Username", "IP", "Success"],
    "Audit_Log": ["Timestamp", "Action", "Section", "Quarter", "FY", "PerformedBy", "Role", "PreviousStatus", "NewStatus", "Comments"],
    "Quarter_Settings": ["FY", "Quarter", "IsLocked", "LockedBy", "LockedDate", "Deadline"],
    "ITA_Certificates": ["FY", "Quarter", "CertifiedBy", "CertifiedDate", "Observations", "Status"],
    "Accounts_MCA": ["Section", "FY", "Quarter", "Month1", "Date1", "Delay1", "Month2", "Date2", "Delay2", "Month3", "Date3", "Delay3", "ExclPct1", "ExclPct2", "ExclPct3"],
    "Accounts_Reconciliation": ["Section", "FY", "Quarter", "Sector", "AmtDue", "AmtReconciled", "PctReconciled"],
    "Accounts_RBD": ["Section", "FY", "Quarter", "MonthDue", "MonthReconciled", "ArrearMonths", "RBI_Amt", "AG_Amt", "NetDiff"],
    "Accounts_Vouchers": ["Section", "FY", "Quarter", "TotalReceived", "TotalDue", "PctChecked"],
    "Accounts_ACBills": ["Section", "FY", "Quarter", "OB", "Addition", "Total", "Clearance", "ClearPct", "CB"],
    "Accounts_UCs": ["Section", "FY", "Quarter", "OB", "Addition", "Total", "Clearance", "ClearPct", "CB"],
    "Suspense_Remittance": ["Section", "FY", "Quarter", "Head", "SubHead", "Dr_Cr", "OB", "Addition", "Total", "Clearance", "ClearPct", "OldCleared", "CB"],
    "Suspense_AnnexC": ["Section", "FY", "Quarter", "Head", "Year", "Dr_Cr", "OB", "Addition", "Total", "Clearance", "ClearPct", "OldCleared", "CB"],
    "MKI_Dates": ["Section", "FY", "Quarter", "Month", "UploadDate"],
    "AnnualAccounts": ["Section", "FY", "Quarter", "Item", "Description", "Date1", "Date2"],
    "LTA_Loans": ["Section", "FY", "Quarter", "Type", "MonthDue", "PostingDone", "Arrear", "PrincipalAmt", "InterestAmt"],
    "TI_Inspection": ["Section", "FY", "Quarter", "TotalPlanned", "InspectedPrevQtr", "PlannedThisQtr", "InspectedThisQtr", "ArrearIf", "OB_IR", "OB_Para", "Add_IR", "Add_Para", "Settled_IR", "Settled_Para", "CB_IR", "CB_Para"],
    "Budget_Review": ["Section", "FY", "Quarter", "DatePassed", "DateReceived", "DateReviewCompleted"],
    "GPF_Summary": ["Section", "FY", "Quarter", "KRA_Item", "OB", "Addition", "Total", "Clearance", "ClearPct", "CB", "Target", "AnnualTarget"],
    "GPF_AnnexH": ["Section", "FY", "Quarter", "CaseType", "SubType", "OB", "Addition", "Total", "Cleared", "ClearPct"],
    "GPF_AnnexI": ["Section", "FY", "Quarter", "Year", "OB", "Addition", "Cleared", "CB"],
    "GPF_AnnexJ": ["Section", "FY", "Quarter", "Type", "Year", "Items_OB", "Amt_OB", "Items_Add", "Amt_Add", "Items_Cleared", "Amt_Cleared", "Items_CB", "Amt_CB"],
    "GPF_AnnexKL": ["Section", "FY", "Quarter", "Annexure", "Year", "OB", "Addition", "Cleared", "CB", "Target", "ClearPct"],
    "GPF_DPF": ["Section", "FY", "Quarter", "OB", "Addition", "Total", "Clearance", "CB"],
    "GPF_Suspense": ["Section", "FY", "Quarter", "Parameter", "CreditAmt", "DebitAmt"],
    "GPF_Online": ["Section", "FY", "Quarter", "Type", "Month", "Subscribers", "EntriesDue", "ReceivedOnline", "PctOnline", "CB"],
    "GPF_Misc": ["Section", "FY", "Quarter", "Field", "Value"],
    "Deposit_PDA": ["Section", "FY", "Quarter", "TotalPDAs", "BalAmt", "Closed", "ClosedAmt", "Opened", "OpenedAmt", "PermissionSought", "Inoperative", "InoperativeAmt"],
    "Complaints_RTI": ["Section", "FY", "Quarter", "Office", "Type", "OB", "Received", "DisposedInTime", "DisposedBeyond", "CB", "DelayReason", "Categories"],
    "Court_Cases": ["Section", "FY", "Quarter", "Office", "CourtType", "OB", "Added", "Decided", "CB", "Categories"],
    "Court_AgeWise": ["Section", "FY", "Quarter", "Office", "AgeGroup", "CaseCount", "Remarks"],
    "Staff_Strength": ["Section", "FY", "Quarter", "Cadre", "SS", "PIP", "VacancyPct"],
    "Staff_GroupWise": ["Section", "FY", "Quarter", "Cadre", "Permanent", "Temporary", "Casual", "AccountsGrp", "FundsGrp", "BhopalBranch", "AdminGrp", "OthersGE", "Total"],
    "VLC_Automation": ["Section", "FY", "Quarter", "AppName", "Database", "OS", "AppPlatform", "QueryPlatform", "DataSource", "DataInput", "OutputForm", "Backup"],
    "VLC_ServiceDisruption": ["Section", "FY", "Quarter", "Function", "ServiceType", "DaysDisrupted", "Reason", "RestorationDate", "ActionTaken"],
    "VLC_ChangeMgmt": ["Section", "FY", "Quarter", "Function", "Description", "CompletionDate", "PendingProposal", "Remarks"],
    "VLC_ARU": ["Section", "FY", "Quarter", "Category", "Numbers"],
    "IFMS_Status": ["Section", "FY", "Quarter", "ModuleName", "Description", "ImplementationStatus", "OnlineData", "InterfaceDeveloped"],
    "IFMS_Online": ["Section", "FY", "Quarter", "SubItem", "Description", "ARU_Total", "ARU_Online", "PctOnline", "ARU_Pending", "Remarks"],
    "Voucher_Details": ["Section", "FY", "Quarter", "VoucherType", "Validated", "TotalReceived", "PctValidated"],
    "Broadsheet_Diff": ["Section", "FY", "Quarter", "HeadOfAccount", "DiffAmount", "ClearanceDuringQtr"],
    "TI_AnnexE_YearWise": ["Section", "FY", "Quarter", "AccYear", "OB_IR", "OB_Para", "Cleared_IR", "Cleared_Para", "CB_IR", "CB_Para", "ActionTaken"],
    "LTA_AnnexD": ["Section", "FY", "Quarter", "Type", "Year", "OB_Items", "OB_Amt", "Cleared_Items", "Cleared_Amt", "CB_Items", "CB_Amt"],
    "Attachments": ["Section", "FY", "Quarter", "FileName", "DriveURL", "FileSize", "UploadTimestamp", "Annexure"]
  };
  
  for (var sheetName in sheets) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.clear();
    sheet.appendRow(sheets[sheetName]);
    
    // Seed default users if it's the Users sheet
    if (sheetName === "Users") {
      sheet.appendRow(["admin", "9438c833b635ac1147c01c715da20215532edab9c4f2ca7e48908ee4ef74dc2b", "System Administrator", "DEVELOPER", "ALL", "Developer", true, new Date(), "system"]);
      sheet.appendRow(["agmp", "2400b17ca054aa108380973e063f53d7176ddfffb54e00a62320579fff5b1fec", "Accountant General MP", "ADMIN", "ALL", "Accountant General", true, new Date(), "system"]);
      sheet.appendRow(["dagaccounts", "2709563864d2194d1ca6f7c05d041bd7bb66cb5edf14efba170fbc53523c27f5", "Dy. Accountant General (Accounts)", "ADMIN", "ALL", "Dy. Accountant General", true, new Date(), "system"]);
      sheet.appendRow(["tm_admin", "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9", "TM Section Admin", "TM_ADMIN", "ALL", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["ita", "6b383ac7c46b2fe4d0db55baa344904ff022baa3afe961e6d188d787859416a6", "ITA Section Officer", "ITA_OFFICER", "ALL", "ITA Officer", true, new Date(), "system"]);
      sheet.appendRow(["aaoita", "c530b0a31b4c6aae7008925535527e36717af2a50d350a35f74516726fcea43c", "Asst. Accounts Officer (ITA)", "ITA_OFFICER", "ALL", "Asst. Accounts Officer (ITA)", true, new Date(), "system"]);
      sheet.appendRow(["sraoita", "c605970cd714fdce369413d0ec5d42a45df50638bd63aceb49b5b7c3155f9cc2", "Sr. Accounts Officer (ITA)", "ITA_OFFICER", "ALL", "Sr. Accounts Officer (ITA)", true, new Date(), "system"]);
      sheet.appendRow(["kra_viewer", "65375049b9e4d7cad6c9ba286fdeb9394b28135a3e84136404cfccfdcc438894", "KRA View-Only Auditor", "VIEWER", "ALL", "View-Only Auditor", true, new Date(), "system"]);
      sheet.appendRow(["book1", "c676562512aa31205cee48429b9e8de069823d3f1e33d87db981dd20d32ef464", "BOOK1 Accountant", "OPERATOR", "BOOK1", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaobook1", "05584fd221251fd46d9f1324644ed8337116e195d8be3ba214feffeece218114", "BOOK1 AAO", "REVIEWER", "BOOK1", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraobook1", "63c6610747052ff994fb94a2dd31dc14ea7680f10574b25e6d98433e705fddfc", "BOOK1 Sr.AO", "SECTION_HEAD", "BOOK1", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["book2", "3cbcece7185ad3f34b6580c5342231d5c407c10b787287d95d8652170d55cbb5", "BOOK2 Accountant", "OPERATOR", "BOOK2", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaobook2", "3afc298764f8868dfa9f9e64b74b8aa35fdf5b7e873fd9219b9d359fd62f6227", "BOOK2 AAO", "REVIEWER", "BOOK2", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraobook2", "adf95313b8f4943e572aa483e18afabd8f85ade7e73e8612d583b8696d5baed7", "BOOK2 Sr.AO", "SECTION_HEAD", "BOOK2", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["rc_cell", "1dd7e8aad2b6880d5d0d5d183e5e212e936cc0e8b7b3c5efd21047d1a5476c85", "RC_CELL Accountant", "OPERATOR", "RC_CELL", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaorc_cell", "2d3448846ee138e75b99a3319e2a4620c90796a54b6eedea1375cb5c884756e3", "RC_CELL AAO", "REVIEWER", "RC_CELL", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraorc_cell", "649017a9735330b9f5d059be5d5cc3b44c4fbd2d512078d88a3961abcfd421e3", "RC_CELL Sr.AO", "SECTION_HEAD", "RC_CELL", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["rbd", "163ee7bd6f7744b41b83cbf40feacf656a8e8e92c2de85b77c3b4685f4708258", "RBD Accountant", "OPERATOR", "RBD", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaorbd", "d8303e3207fd1ed0d30c9e3ba251f969201a5a6e141cfc0d7f29ad1c3bed36c1", "RBD AAO", "REVIEWER", "RBD", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraorbd", "5e7dbc90e4afb1e6ff6ec6c14c26c98c845d5e34236e5557ea84207d128bdfd9", "RBD Sr.AO", "SECTION_HEAD", "RBD", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["cv_cell", "0b6816dca09d7a3b6bade99c7d557c48ff24173c54cac2d005b37f065b67f8d5", "CV_CELL Accountant", "OPERATOR", "CV_CELL", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaocv_cell", "abe573a1f767741e23c092983250a9b980666a0706780843b60cca31d640002b", "CV_CELL AAO", "REVIEWER", "CV_CELL", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraocv_cell", "d90ad22578ab2838b3b48d5fe26942375032a0dd242ee0b7571b8ac33a7686f1", "CV_CELL Sr.AO", "SECTION_HEAD", "CV_CELL", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["tm", "6dee04b73e8d3336654bca37a2d58c001f0bfc572f4bf0d46c38fc08002d4da7", "TM Accountant", "OPERATOR", "TM", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaotm", "bda65c3921f0253dc067a7b58f56317a12be2c857ef6c8dbcd29073ed0fd73c8", "TM AAO", "REVIEWER", "TM", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraotm", "f8adad129581ce843de727ca670eb83dd115ef36ce95c42257124a5a3933ffc2", "TM Sr.AO", "SECTION_HEAD", "TM", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["acd", "b647eb1fbb0e2e29c14dbe1d3ed6d706eb5ae292e01eab9f1c52e003a9e08a02", "ACD Accountant", "OPERATOR", "ACD", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaoacd", "fcbe3f270d6e5be42781fd40f5c663e38c36499ddd1cc3ba110bb6ba326aa34a", "ACD AAO", "REVIEWER", "ACD", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraoacd", "ce3e210369fac672a90f5b4bdb47c6839be69c52361a9812ead15ebbdd08e7ad", "ACD Sr.AO", "SECTION_HEAD", "ACD", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["ct", "f8dcb34308e5c69f46a33cafa62bd922476b18c78a114eb144e34a5e374f7d60", "CT Accountant", "OPERATOR", "CT", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaoct", "de2a3ca6c96b7f9001a80c3ed755420de0c8a455311badbb909f4ec603a71ac0", "CT AAO", "REVIEWER", "CT", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraoct", "ffbb41a3d5351ef07216e8969120dd3dc9c224380d3a42e4a1ec0823d507c82d", "CT Sr.AO", "SECTION_HEAD", "CT", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["fc", "8312c5d48baec3e2cb31e1402d2eedd77ff6dc19f9f64895ca7fe6e23a73a636", "FC Accountant", "OPERATOR", "FC", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaofc", "b8f9ba7cc122ebeb280c8223bb77b7e87d8e4ab10db938fc545063112d9074b4", "FC AAO", "REVIEWER", "FC", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraofc", "504432d75c39ef5db19cdc79cd18baec9d0506bc690262c5a4d71c86f325d395", "FC Sr.AO", "SECTION_HEAD", "FC", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["bhopal", "e261391dc6be9ec1594c20f8385f213cfd26007e0c6ca56a92c7bee5d874bc01", "BHOPAL Accountant", "OPERATOR", "BHOPAL", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaobhopal", "2c727193faaf00f59fcb2273d2b37ba5b361c3f878449003e7d217f613646819", "BHOPAL AAO", "REVIEWER", "BHOPAL", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraobhopal", "6b8983e6678bd1c25fdb04ceceb6841cbbcfb85c060a1e902b5e6f53c49b3cb2", "BHOPAL Sr.AO", "SECTION_HEAD", "BHOPAL", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["loan", "472bbf14923e2e7cefd8529825c401e8d1a2937b96dd697a6d1c75c53e6cca3a", "LOAN Accountant", "OPERATOR", "LOAN", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaoloan", "8bad1d5133560731a25b595e41b1e80c1130da0a9a3fe2b9d31f355edfa7c05a", "LOAN AAO", "REVIEWER", "LOAN", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraoloan", "12d8b97b7fff0fd365ff3b761d574dae1eeae04628346b9a18e6696b52e28ad5", "LOAN Sr.AO", "SECTION_HEAD", "LOAN", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["ti_cell", "15dc420b2cf3f6211cf461c75e0432f77f08eba2196ce25af3312574735f8dc7", "TI_CELL Accountant", "OPERATOR", "TI_CELL", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaoti_cell", "f06db34267d3be3967a9e73908b532cb5f7b86fa0d05a7dfd1e7977c51e5558a", "TI_CELL AAO", "REVIEWER", "TI_CELL", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraoti_cell", "610f9781829ab22560503669c990fa6096cee3b15b7d7a1ba67d7cbd8518daf2", "TI_CELL Sr.AO", "SECTION_HEAD", "TI_CELL", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["fund", "639f78fb7729d09dc6066a6dd81997572d2413c50b703edd5b378166b5466b2d", "FUND Accountant", "OPERATOR", "FUND", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaofund", "ca4c3504049e971b2bd783b5a93863117514087106b24bde9363e04dbbbe1728", "FUND AAO", "REVIEWER", "FUND", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraofund", "94735dd3f674c6a1868296f161aa7150ab997b3bac6d14ffc3b76a6b0ef66cd7", "FUND Sr.AO", "SECTION_HEAD", "FUND", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["deposit", "c3b9fb78a452ce2fc90cff1608510235503e3b727683b71c7fefee54198bad63", "DEPOSIT Accountant", "OPERATOR", "DEPOSIT", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaodeposit", "f0e893d451f653f36099fd1b590f7272e5df1fcc1f0b671477e475a00d505989", "DEPOSIT AAO", "REVIEWER", "DEPOSIT", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraodeposit", "1424695d2da15b66dc749121fac641b4e625f8a215a5233c6e1fad2684f4ed27", "DEPOSIT Sr.AO", "SECTION_HEAD", "DEPOSIT", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sec_admin", "380b2a8ed7f201083e9b110bce11fefb20c29ed7eb372a76f272a2e4125b742b", "ADMIN Accountant", "OPERATOR", "ADMIN", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaosec_admin", "3ec0f0c0ae2f845fa8aebc5aeec2aa46d84a7e7811f0556f2f01f2f01f01f01f", "ADMIN AAO", "REVIEWER", "ADMIN", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraosec_admin", "5ec0f0c0ae2f845fa8aebc5aeec2aa46d84a7e7811f0556f2f01f2f01f01f01f", "ADMIN Sr.AO", "SECTION_HEAD", "ADMIN", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["vlc", "51c80ba0161295bb0ba04257e74b12e29b28234f79253400f0e803db107c3b84", "VLC Accountant", "OPERATOR", "VLC", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaovlc", "1078a2292d7e9bdca9da18e56f354398dbadc70d030ce2cb65bd238f4439f021", "VLC AAO", "REVIEWER", "VLC", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraovlc", "be9eaf1db84b0f9f3240dec9b65b7b822b1ba9590902f9c802c85c6cf39a103b", "VLC Sr.AO", "SECTION_HEAD", "VLC", "Sr. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["report", "845e91831319e89c4d656bdb80c278ac09a7230d61e5dfd2e1b1fbb436ac8917", "REPORT Accountant", "OPERATOR", "REPORT", "Accountant", true, new Date(), "system"]);
      sheet.appendRow(["aaoreport", "9d113363860db51fb423a453fd98d517d17e09430946af2ae1a0e06552a76ad3", "REPORT AAO", "REVIEWER", "REPORT", "Asst. Accounts Officer", true, new Date(), "system"]);
      sheet.appendRow(["sraoreport", "1af8ad2efbe6cbb3ca889d3940f01727737aae5ce315aaf81a460308e0882c89", "REPORT Sr.AO", "SECTION_HEAD", "REPORT", "Sr. Accounts Officer", true, new Date(), "system"]);
    }
  }
}
