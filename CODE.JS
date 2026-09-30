/**
 * ============================================================================
 * COMPLETE FACULTY APPOINTMENT, PAYROLL, BIN RECYCLER, BULK UPLOAD & WEB PORTAL
 * ============================================================================
 * Column Structure (Sheet1 & BIN RECORDS):
 * Col E (5) : UID
 * Col F (6) : Teacher Name
 * Col G (7) : Gender
 * Col H (8) : Base Salary
 * Col I (9) : Department
 * Col J (10): Joining Date
 * Col K (11): Address             <-- NEW COLUMN
 * Col L (12): Bonus
 * Col M (13): Phone
 * Col N (14): Email               <-- Last 15 characters shown
 * Col O (15): Status
 * Col P (16): Deleted Timestamp   <-- in BIN RECORDS
 * 
 * CONTACT US:
 * Col A (1) : Request ID
 * Col B (2) : Timestamp
 * Col C (3) : Faculty UID
 * Col D (4) : Faculty Name
 * Col E (5) : Phone
 * Col F (6) : Email
 * Col G (7) : Request Type
 * Col H (8) : Description
 * Col I (9) : Attachment Link
 * Col J (10): Status (Pending / In Progress / Resolved / Rejected)
 */

var SHEET_MAIN = "Sheet1";
var SHEET_BIN = "BIN RECORDS";
var SHEET_CONTACT = "CONTACT US";

var COL_UID = 5;      // E
var COL_NAME = 6;     // F
var COL_GENDER = 7;   // G
var COL_SALARY = 8;   // H
var COL_DEPT = 9;     // I
var COL_DATE = 10;    // J
var COL_ADDRESS = 11; // K
var COL_BONUS = 12;   // L
var COL_PHONE = 13;   // M
var COL_EMAIL = 14;   // N
var COL_STATUS = 15;  // O
var COL_BIN_TIME = 16;// P
var HEADER_ROW = 2;

var CU_REQ_ID = 1;    // A
var CU_TIME = 2;      // B
var CU_UID = 3;       // C
var CU_NAME = 4;      // D
var CU_PHONE = 5;     // E
var CU_EMAIL = 6;     // F
var CU_TYPE = 7;      // G
var CU_DESC = 8;      // H
var CU_ATTACH = 9;    // I
var CU_STATUS = 10;   // J

var OTP_CACHE = CacheService.getScriptCache();

/**
 * Serves the Web Portal HTML.
 */
function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle("Lovely Professional University - Faculty Portal")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  } catch (err) {
    return HtmlService.createHtmlOutput(getPortalHtml())
      .setTitle("Lovely Professional University - Faculty Portal")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  }
}

/**
 * Handles internal RPC requests from google.script.run (CORS-free).
 */
function handleApiRequest(action, payloadJson) {
  var request = JSON.parse(payloadJson || "{}");
  request.action = action;
  return processApiAction(request);
}

/**
 * Universal POST API Router for external/fetch requests.
 */
function doPost(e) {
  try {
    var contents = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var request = JSON.parse(contents);
    var responseData = processApiAction(request);

    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Central API Action Router
 */
function processApiAction(request) {
  var action = request.action;
  if (action === "apiAdminLogin") return apiAdminLogin(request.username, request.password);
  if (action === "apiStaffRequestOtp") return apiStaffRequestOtp(request.identifier, request.joiningDate);
  if (action === "apiStaffVerifyOtp") return apiStaffVerifyOtp(request.uid, request.enteredOtp);
  if (action === "apiStaffUpdateContact") return apiStaffUpdateContact(request.uid, request.phone, request.email);
  if (action === "apiStaffSubmitRequest") return apiStaffSubmitRequest(request.data);
  if (action === "apiAdminGetAllData") return apiAdminGetAllData();
  if (action === "apiAdminAddRecord") return apiAdminAddRecord(request.record);
  if (action === "apiAdminUpdateRecord") return apiAdminUpdateRecord(request.record);
  if (action === "apiAdminBulkUpload") return apiAdminBulkUpload(request.records);
  if (action === "apiAdminMoveToBin") return apiAdminMoveToBin(request.uid);
  if (action === "apiAdminRestoreRecord") return apiAdminRestoreRecord(request.uid);
  if (action === "apiAdminUpdateRequestStatus") return apiAdminUpdateRequestStatus(request.reqId, request.newStatus);
  return { success: false, message: "Unknown action: " + action };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏫 Faculty & Bin Portal")
    .addItem("⚡ 1-Click ACTIVATE Auto Triggers", "installMyTrigger")
    .addItem("🧪 Send Instant Test Email & PDF to Me", "sendTestEmailNow")
    .addItem("📧 Dispatch Statements to All Active Teachers", "sendStatementsForExistingRecords")
    .addItem("📜 Send 2-Page Offer Letter for Selected Row", "sendWelcomeOfferForSelectedRow")
    .addItem("🗑️ Purge Expired Bin Records (>30s)", "purgeExpiredBinRecords")
    .addToUi();
}

/**
 * 1-Click Trigger Installer
 */
function installMyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  ScriptApp.newTrigger("handleSheetEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast("Auto-Triggers are ACTIVE across Sheet1, Bin & Contact Us!", "Triggers Live", 6);
  Logger.log("Installable onEdit trigger created successfully.");
}

/**
 * --------------------------------------------------------------------------
 * BULK UPLOAD ENGINE (AUTO COLUMN MAPPING & ZERO DUPLICATES BY NAME)
 * --------------------------------------------------------------------------
 */
function apiAdminBulkUpload(records) {
  if (!records || !Array.isArray(records) || records.length === 0) {
    return { success: false, message: "No records found in uploaded file." };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);
  var lastRow = sheet.getLastRow();

  var nameToRowMap = {};
  if (lastRow > HEADER_ROW) {
    var existingNames = sheet.getRange(HEADER_ROW + 1, COL_NAME, lastRow - HEADER_ROW, 1).getValues();
    for (var r = 0; r < existingNames.length; r++) {
      var n = String(existingNames[r][0] || "").trim().toLowerCase().replace(/\s+/g, ' ');
      if (n) {
        nameToRowMap[n] = HEADER_ROW + 1 + r;
      }
    }
  }

  var addedCount = 0;
  var updatedCount = 0;

  for (var i = 0; i < records.length; i++) {
    var item = records[i];
    var rawName = String(item.name || "").trim();
    if (!rawName) continue;

    var nameKey = rawName.toLowerCase().replace(/\s+/g, ' ');
    var gender = item.gender ? String(item.gender).trim() : "Female";
    var salary = Number(String(item.salary || 0).replace(/[^0-9.]/g, "")) || 50000;
    var dept = item.dept ? String(item.dept).trim() : "General Faculty";
    var jDate = item.joiningDate ? String(item.joiningDate).trim() : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var address = item.address ? String(item.address).trim() : "";
    var bonus = calculateBonusFromJoiningDate(jDate);
    var phone = String(item.phone || "").trim();
    var email = String(item.email || "").trim();

    if (nameToRowMap.hasOwnProperty(nameKey)) {
      var targetRow = nameToRowMap[nameKey];
      var oldSalary = sheet.getRange(targetRow, COL_SALARY).getValue();
      var oldDept = sheet.getRange(targetRow, COL_DEPT).getValue();
      var oldAddress = sheet.getRange(targetRow, COL_ADDRESS).getValue();
      var oldPhone = sheet.getRange(targetRow, COL_PHONE).getValue();

      sheet.getRange(targetRow, COL_GENDER).setValue(gender);
      sheet.getRange(targetRow, COL_SALARY).setValue(salary);
      sheet.getRange(targetRow, COL_DEPT).setValue(dept);
      sheet.getRange(targetRow, COL_DATE).setValue(jDate);
      sheet.getRange(targetRow, COL_ADDRESS).setValue(address);
      sheet.getRange(targetRow, COL_BONUS).setValue(bonus);
      sheet.getRange(targetRow, COL_PHONE).setValue(phone);
      sheet.getRange(targetRow, COL_EMAIL).setValue(email);

      if (email) {
        var oldSummary = "Salary: ₹" + Number(oldSalary).toLocaleString('en-IN') + ", Dept: " + oldDept + ", Address: " + oldAddress + ", Phone: " + oldPhone;
        var newSummary = "Salary: ₹" + Number(salary).toLocaleString('en-IN') + ", Dept: " + dept + ", Address: " + address + ", Phone: " + phone;
        sendStatementForRow(sheet, targetRow, "Bulk Sync Record Update", oldSummary, newSummary);
      }
      updatedCount++;
    } else {
      var newUid = generateUnique6DigitUid(sheet, COL_UID);
      var nextRow = getNextAvailableRow(sheet, COL_UID, HEADER_ROW + 1);

      var newRowData = [
        newUid,
        rawName,
        gender,
        salary,
        dept,
        jDate,
        address,
        bonus,
        phone,
        email,
        "Active"
      ];

      sheet.getRange(nextRow, COL_UID, 1, newRowData.length).setValues([newRowData]);
      nameToRowMap[nameKey] = nextRow;

      if (email) {
        sendWelcomeOfferLetter(sheet, nextRow);
      }
      addedCount++;
    }
  }

  return {
    success: true,
    addedCount: addedCount,
    updatedCount: updatedCount,
    totalProcessed: addedCount + updatedCount,
    message: "Bulk upload complete! Added: " + addedCount + " new faculty, Updated: " + updatedCount + " existing records (No duplicates created)."
  };
}

/**
 * --------------------------------------------------------------------------
 * AUTHENTICATION (ADMIN & STAFF WITH JOINING DATE & 2-MIN OTP)
 * --------------------------------------------------------------------------
 */

function apiAdminLogin(username, password) {
  if (username === "001" && password === "001") {
    return { success: true, message: "Admin authenticated." };
  }
  return { success: false, message: "Invalid Admin User ID or Password." };
}

function normalizeDateStr(d) {
  if (!d) return "";
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var parsed = new Date(d);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(d).trim();
}

function apiStaffRequestOtp(identifier, joiningDate) {
  var cleanId = String(identifier || "").trim();
  var inputDate = normalizeDateStr(joiningDate);

  if (!cleanId) return { success: false, message: "Please enter your UID or Mobile number." };
  if (!inputDate) return { success: false, message: "Please select your Date of Joining." };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var binSheet = ss.getSheetByName(SHEET_BIN);

  if (binSheet && binSheet.getLastRow() > HEADER_ROW) {
    var binData = binSheet.getRange(HEADER_ROW + 1, COL_UID, binSheet.getLastRow() - HEADER_ROW, 11).getValues();
    for (var b = 0; b < binData.length; b++) {
      var bUid = String(binData[b][0]).trim();
      var bPhone = String(binData[b][8]).replace(/\D/g, "");
      var cleanInput = cleanId.replace(/\D/g, "");
      if (bUid === cleanId || (cleanInput && bPhone.endsWith(cleanInput))) {
        return { success: false, message: "No record found. Profile is marked inactive/deleted. Please contact Administrator." };
      }
    }
  }

  var lastRow = mainSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { success: false, message: "No record found. Please contact Administrator." };

  var data = mainSheet.getRange(HEADER_ROW + 1, COL_UID, lastRow - HEADER_ROW, 11).getValues();
  var foundTeacher = null;

  for (var i = 0; i < data.length; i++) {
    var tUid = String(data[i][0]).trim();
    var tPhone = String(data[i][8]).replace(/\D/g, "");
    var tEmail = String(data[i][9]).trim();
    var tStatus = String(data[i][10]).trim().toUpperCase();
    var tDate = normalizeDateStr(data[i][5]);
    var cleanInputPhone = cleanId.replace(/\D/g, "");

    if (tUid === cleanId || (cleanInputPhone.length >= 4 && tPhone.endsWith(cleanInputPhone))) {
      if (tStatus !== "ACTIVE") {
        return { success: false, message: "No record found. Account is marked closed. Please contact Administrator." };
      }
      if (tDate !== inputDate) {
        return { success: false, message: "Incorrect Date of Joining. Verification failed." };
      }

      foundTeacher = {
        row: HEADER_ROW + 1 + i,
        uid: tUid,
        name: data[i][1],
        email: tEmail,
        phone: data[i][8]
      };
      break;
    }
  }

  if (!foundTeacher) return { success: false, message: "No record found. Please contact Administrator." };
  if (!foundTeacher.email) return { success: false, message: "No registered email found for this account. Please contact Administrator." };

  var otp = Math.floor(100000 + Math.random() * 900000).toString();
  OTP_CACHE.put("OTP_" + foundTeacher.uid, otp, 120);

  var emailBody = "<div style='font-family:Arial,sans-serif; padding:20px; border:1px solid #e2e8f0; border-radius:10px; max-width:500px;'>" +
    "<h2 style='color:#4338ca; margin-top:0;'>Faculty Portal OTP Verification</h2>" +
    "<p>Hello <strong>" + foundTeacher.name + "</strong>,</p>" +
    "<p>Your one-time password (OTP) for accessing your faculty statement is:</p>" +
    "<div style='background:#f1f5f9; padding:12px; font-size:26px; font-weight:800; letter-spacing:5px; text-align:center; color:#1e1b4b; border-radius:6px; margin:15px 0;'>" + otp + "</div>" +
    "<p style='color:#dc2626; font-size:12.5px; font-weight:600;'>⏱️ This OTP is valid for 2 MINUTES. Do not share this code with anyone.</p>" +
    "<br><div style='background:#fef2f2; border:1px solid #fecaca; padding:10px; border-radius:6px; font-size:11px; color:#991b1b; text-align:center; font-weight:700;'>" +
    "⚠️ THIS IS A SYSTEM GENERATED NOTIFICATION EMAIL. PLEASE DO NOT REPLY. IF RECEIVED BY MISTAKE, PLEASE DESTROY THIS EMAIL IMMEDIATELY." +
    "</div></div>";

  MailApp.sendEmail({
    to: foundTeacher.email,
    subject: "Your 2-Minute Faculty Portal OTP: " + otp,
    htmlBody: emailBody
  });

  var maskedEmail = foundTeacher.email.length > 15 
    ? "***************" + foundTeacher.email.slice(-15) 
    : foundTeacher.email;

  return { success: true, uid: foundTeacher.uid, message: "OTP successfully dispatched to registered email (" + maskedEmail + ")." };
}

function apiStaffVerifyOtp(uid, enteredOtp) {
  var storedOtp = OTP_CACHE.get("OTP_" + uid);
  if (!storedOtp) return { success: false, message: "OTP has expired after 2 minutes. Please click 'Resend OTP' to get a new code." };
  if (storedOtp !== String(enteredOtp).trim()) return { success: false, message: "Incorrect OTP entered. Please verify and try again." };

  OTP_CACHE.remove("OTP_" + uid);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { success: false, message: "Unable to retrieve record details." };
  var data = sheet.getRange(HEADER_ROW + 1, COL_UID, lastRow - HEADER_ROW, 11).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === uid) {
      var salary = Number(data[i][3]) || 0;
      var bonus = Number(data[i][7]) || 0;
      var joiningDate = data[i][5] instanceof Date 
        ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "dd/MM/yyyy") 
        : data[i][5];

      var address = String(data[i][6] || "N/A");
      var rawPhone = String(data[i][8] || "");
      var rawEmail = String(data[i][9] || "");

      return {
        success: true,
        record: {
          uid: data[i][0],
          name: data[i][1],
          gender: data[i][2],
          salary: salary,
          dept: data[i][4],
          joiningDate: joiningDate,
          address: address,
          bonus: bonus,
          total: salary + bonus,
          phone: rawPhone,
          email: rawEmail,
          maskedPhone: rawPhone.length > 4 ? "***************" + rawPhone.slice(-4) : rawPhone,
          maskedEmail: rawEmail.length > 15 ? "***************" + rawEmail.slice(-15) : rawEmail,
          status: data[i][10]
        }
      };
    }
  }

  return { success: false, message: "Unable to retrieve record details." };
}

function apiStaffUpdateContact(uid, newPhone, newEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { success: false, message: "Record not found." };
  var data = sheet.getRange(HEADER_ROW + 1, COL_UID, lastRow - HEADER_ROW, 11).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(uid).trim()) {
      var row = HEADER_ROW + 1 + i;
      var oldPhone = String(data[i][8] || "");
      var oldEmail = String(data[i][9] || "");

      sheet.getRange(row, COL_PHONE).setValue(newPhone);
      sheet.getRange(row, COL_EMAIL).setValue(newEmail);

      var changeDesc = "Contact Credentials Updated: Phone & Email";
      var oldValText = "Phone: " + oldPhone + " | Email: " + oldEmail;
      var newValText = "Phone: " + newPhone + " | Email: " + newEmail;

      sendStatementForRow(sheet, row, changeDesc, oldValText, newValText);
      return { success: true, message: "Contact information updated successfully!" };
    }
  }
  return { success: false, message: "Record not found." };
}

function apiStaffSubmitRequest(req) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cuSheet = ss.getSheetByName(SHEET_CONTACT);
    if (!cuSheet) return { success: false, message: "CONTACT US sheet not found." };

    var fileUrl = "No Attachment";
    if (req.fileBase64 && req.fileName) {
      var bytes = Utilities.base64Decode(req.fileBase64.split(",")[1] || req.fileBase64);
      var blob = Utilities.newBlob(bytes, req.fileType, req.fileName);
      var file = DriveApp.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    var nextRow = cuSheet.getLastRow() + 1;
    var reqId = "REQ-" + Utilities.formatString("%03d", nextRow - 1);
    var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    var rowData = [
      reqId,
      timestampStr,
      req.uid,
      req.name,
      req.phone,
      req.email,
      req.type,
      req.description,
      fileUrl,
      "Pending"
    ];

    cuSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    var confirmEmailHtml = "<div style='font-family:Arial,sans-serif; padding:20px; border:1px solid #e2e8f0; border-radius:10px; max-width:600px;'>" +
      "<h2 style='color:#4338ca;'>Support Request Received: " + reqId + "</h2>" +
      "<p>Dear <strong>" + req.name + "</strong>,</p>" +
      "<p>Your support ticket has been submitted to the Academic Administration:</p>" +
      "<table style='width:100%; border-collapse:collapse; margin:15px 0; font-size:13px;'>" +
      "<tr style='background:#f8fafc;'><td style='padding:8px; border:1px solid #e2e8f0;'>Request ID</td><td style='padding:8px; border:1px solid #e2e8f0;'><strong>" + reqId + "</strong></td></tr>" +
      "<tr><td style='padding:8px; border:1px solid #e2e8f0;'>Request Type</td><td style='padding:8px; border:1px solid #e2e8f0;'>" + req.type + "</td></tr>" +
      "<tr style='background:#f8fafc;'><td style='padding:8px; border:1px solid #e2e8f0;'>Description</td><td style='padding:8px; border:1px solid #e2e8f0;'>" + req.description + "</td></tr>" +
      "<tr><td style='padding:8px; border:1px solid #e2e8f0;'>Attachment</td><td style='padding:8px; border:1px solid #e2e8f0;'><a href='" + fileUrl + "'>View Uploaded File</a></td></tr>" +
      "<tr style='background:#f8fafc;'><td style='padding:8px; border:1px solid #e2e8f0;'>Status</td><td style='padding:8px; border:1px solid #e2e8f0;'><strong style='color:#d97706;'>Pending Review</strong></td></tr>" +
      "</table>" +
      "<div style='background:#fef2f2; border:1px solid #fecaca; padding:10px; border-radius:6px; font-size:11px; color:#991b1b; text-align:center; font-weight:700; margin-top:20px;'>" +
      "THIS IS A SYSTEM GENERATED NOTIFICATION EMAIL. PLEASE DO NOT REPLY. IF RECEIVED BY MISTAKE, PLEASE DESTROY THIS EMAIL IMMEDIATELY." +
      "</div></div>";

    MailApp.sendEmail({
      to: req.email,
      subject: "Support Request Logged [" + reqId + "] - " + req.type,
      htmlBody: confirmEmailHtml
    });

    ss.toast("Support Request " + reqId + " submitted successfully!", "✅ Ticket Logged", 5);
    return { success: true, reqId: reqId, message: "Request " + reqId + " submitted successfully!" };
  } catch (err) {
    return { success: false, message: "Error submitting request: " + err.toString() };
  }
}

/**
 * --------------------------------------------------------------------------
 * ADMIN DATA RETRIEVAL & ACTIONS
 * --------------------------------------------------------------------------
 */

function apiAdminGetAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var binSheet = ss.getSheetByName(SHEET_BIN);
  var cuSheet = ss.getSheetByName(SHEET_CONTACT);

  var activeRecords = [];
  var binRecords = [];
  var contactRequests = [];

  if (mainSheet && mainSheet.getLastRow() > HEADER_ROW) {
    var mainData = mainSheet.getRange(HEADER_ROW + 1, COL_UID, mainSheet.getLastRow() - HEADER_ROW, 11).getValues();
    for (var i = 0; i < mainData.length; i++) {
      if (mainData[i][0]) {
        var jDate = mainData[i][5] instanceof Date 
          ? Utilities.formatDate(mainData[i][5], Session.getScriptTimeZone(), "yyyy-MM-dd") 
          : mainData[i][5];
        activeRecords.push({
          row: HEADER_ROW + 1 + i,
          uid: mainData[i][0],
          name: mainData[i][1],
          gender: mainData[i][2],
          salary: Number(mainData[i][3]) || 0,
          dept: mainData[i][4],
          joiningDate: jDate,
          address: mainData[i][6] || "",
          bonus: Number(mainData[i][7]) || 0,
          phone: mainData[i][8],
          email: mainData[i][9],
          status: mainData[i][10]
        });
      }
    }
  }

  if (binSheet && binSheet.getLastRow() > HEADER_ROW) {
    var binData = binSheet.getRange(HEADER_ROW + 1, COL_UID, binSheet.getLastRow() - HEADER_ROW, 12).getValues();
    for (var b = 0; b < binData.length; b++) {
      if (binData[b][0]) {
        var bDate = binData[b][5] instanceof Date 
          ? Utilities.formatDate(binData[b][5], Session.getScriptTimeZone(), "yyyy-MM-dd") 
          : binData[b][5];
        var bTime = binData[b][11] instanceof Date 
          ? Utilities.formatDate(binData[b][11], Session.getScriptTimeZone(), "hh:mm:ss a") 
          : "Recently";
        binRecords.push({
          row: HEADER_ROW + 1 + b,
          uid: binData[b][0],
          name: binData[b][1],
          dept: binData[b][4],
          joiningDate: bDate,
          address: binData[b][6] || "",
          salary: Number(binData[b][3]) || 0,
          bonus: Number(binData[b][7]) || 0,
          status: binData[b][10],
          deletedAt: bTime
        });
      }
    }
  }

  if (cuSheet && cuSheet.getLastRow() > 1) {
    var cuData = cuSheet.getRange(2, 1, cuSheet.getLastRow() - 1, 10).getValues();
    for (var c = 0; c < cuData.length; c++) {
      if (cuData[c][0]) {
        var reqTime = cuData[c][1] instanceof Date 
          ? Utilities.formatDate(cuData[c][1], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") 
          : cuData[c][1];
        contactRequests.push({
          reqId: cuData[c][0],
          timestamp: reqTime,
          uid: cuData[c][2],
          name: cuData[c][3],
          phone: cuData[c][4],
          email: cuData[c][5],
          type: cuData[c][6],
          desc: cuData[c][7],
          attachUrl: cuData[c][8],
          status: cuData[c][9]
        });
      }
    }
  }

  return { active: activeRecords, bin: binRecords, requests: contactRequests };
}

function apiAdminAddRecord(record) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);

  var uniqueUid = generateUnique6DigitUid(sheet, COL_UID);
  var calculatedBonus = calculateBonusFromJoiningDate(record.joiningDate);
  var nextRow = getNextAvailableRow(sheet, COL_UID, HEADER_ROW + 1);

  var newRowData = [
    uniqueUid,
    record.name,
    record.gender,
    Number(record.salary) || 0,
    record.dept,
    record.joiningDate,
    record.address || "",
    calculatedBonus,
    record.phone,
    record.email,
    "Active"
  ];

  sheet.getRange(nextRow, COL_UID, 1, newRowData.length).setValues([newRowData]);
  sendWelcomeOfferLetter(sheet, nextRow);
  return { success: true, uid: uniqueUid };
}

function apiAdminUpdateRecord(record) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { success: false, message: "Record not found." };
  var data = sheet.getRange(HEADER_ROW + 1, COL_UID, lastRow - HEADER_ROW, 1).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(record.uid).trim()) {
      var row = HEADER_ROW + 1 + i;
      var calculatedBonus = calculateBonusFromJoiningDate(record.joiningDate);

      var oldSalary = sheet.getRange(row, COL_SALARY).getValue();
      var oldAddress = sheet.getRange(row, COL_ADDRESS).getValue();
      var oldPhone = sheet.getRange(row, COL_PHONE).getValue();
      var oldEmail = sheet.getRange(row, COL_EMAIL).getValue();

      sheet.getRange(row, COL_NAME).setValue(record.name);
      sheet.getRange(row, COL_GENDER).setValue(record.gender);
      sheet.getRange(row, COL_SALARY).setValue(Number(record.salary) || 0);
      sheet.getRange(row, COL_DEPT).setValue(record.dept);
      sheet.getRange(row, COL_DATE).setValue(record.joiningDate);
      sheet.getRange(row, COL_ADDRESS).setValue(record.address || "");
      sheet.getRange(row, COL_BONUS).setValue(calculatedBonus);
      sheet.getRange(row, COL_PHONE).setValue(record.phone);
      sheet.getRange(row, COL_EMAIL).setValue(record.email);

      var oldSummary = "Salary: ₹" + Number(oldSalary).toLocaleString('en-IN') + ", Address: " + oldAddress + ", Phone: " + oldPhone + ", Email: " + oldEmail;
      var newSummary = "Salary: ₹" + Number(record.salary).toLocaleString('en-IN') + ", Address: " + record.address + ", Phone: " + record.phone + ", Email: " + record.email;

      sendStatementForRow(sheet, row, "Portal Profile Revision", oldSummary, newSummary);
      return { success: true };
    }
  }
  return { success: false, message: "Record not found." };
}

function apiAdminMoveToBin(uid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { success: false, message: "Record not found." };
  var data = sheet.getRange(HEADER_ROW + 1, COL_UID, lastRow - HEADER_ROW, 1).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(uid).trim()) {
      var row = HEADER_ROW + 1 + i;
      processDeletionAndMoveToBin(row, "CLOSED");
      return { success: true };
    }
  }
  return { success: false, message: "Record not found." };
}

function apiAdminRestoreRecord(uid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var binSheet = ss.getSheetByName(SHEET_BIN);
  if (!binSheet) return { success: false, message: "Bin sheet not found." };

  var lastRow = binSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { success: false, message: "Record not found in Bin." };
  var data = binSheet.getRange(HEADER_ROW + 1, COL_UID, lastRow - HEADER_ROW, 1).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(uid).trim()) {
      var row = HEADER_ROW + 1 + i;
      restoreRowToMainSheet(row);
      return { success: true };
    }
  }
  return { success: false, message: "Record not found in Bin." };
}

function apiAdminUpdateRequestStatus(reqId, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cuSheet = ss.getSheetByName(SHEET_CONTACT);
  if (!cuSheet) return { success: false, message: "CONTACT US sheet not found." };

  var lastRow = cuSheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: "Request ID not found." };
  var data = cuSheet.getRange(2, CU_REQ_ID, lastRow - 1, 1).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(reqId).trim()) {
      var row = 2 + i;
      cuSheet.getRange(row, CU_STATUS).setValue(newStatus);
      processContactUsStatusChange(row, newStatus);
      return { success: true };
    }
  }
  return { success: false, message: "Request ID not found." };
}

/**
 * --------------------------------------------------------------------------
 * AUTOMATED SPREADSHEET EVENT HANDLERS
 * --------------------------------------------------------------------------
 */

function handleSheetEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  // 1. Actions on CONTACT US Sheet
  if (sheetName === SHEET_CONTACT && row > 1 && col === CU_STATUS) {
    var cuStatus = String(e.value || "").trim();
    processContactUsStatusChange(row, cuStatus);
    return;
  }

  if (row <= HEADER_ROW) return;

  // 2. Actions on BIN RECORDS
  if (sheetName === SHEET_BIN && col === COL_STATUS) {
    if (String(e.value || "").trim().toUpperCase() === "RESTORE") {
      restoreRowToMainSheet(row);
    }
    return;
  }

  if (sheetName !== SHEET_MAIN) return;

  // 3. Immutable UID Check
  if (col === COL_UID && e.oldValue && e.oldValue !== e.value) {
    e.range.setValue(e.oldValue);
    SpreadsheetApp.getActiveSpreadsheet().toast("UID is locked and cannot be modified.", "⚠️ Security Protection", 3);
    return;
  }

  // 4. Auto UID Generation on Name Entry
  if (col === COL_NAME && e.value && !sheet.getRange(row, COL_UID).getValue()) {
    var newUid = generateUnique6DigitUid(sheet, COL_UID);
    sheet.getRange(row, COL_UID).setValue(newUid);
    SpreadsheetApp.getActiveSpreadsheet().toast("Generated UID: " + newUid, "✅ New Record", 3);
  }

  // 5. Auto Bonus on Joining Date Entry/Edit
  if (col === COL_DATE && e.value) {
    var calculatedBonus = calculateBonusFromJoiningDate(e.value);
    sheet.getRange(row, COL_BONUS).setValue(calculatedBonus);
    SpreadsheetApp.getActiveSpreadsheet().toast("Auto-calculated Bonus: ₹" + calculatedBonus.toLocaleString('en-IN'), "💰 Bonus Computed", 4);
  }

  // 6. CLOSED / INACTIVE / DELETED: Move to Bin with 30s Grace Period
  if (col === COL_STATUS) {
    var status = String(e.value || "").trim().toUpperCase();
    if (status === "CLOSED" || status === "INACTIVE" || status === "DELETED") {
      processDeletionAndMoveToBin(row, status);
      return;
    }
  }

  // 7. Field Modification Emails (Cols F-N)
  if (col >= COL_NAME && col <= COL_EMAIL) {
    if (col === COL_EMAIL && (!e.oldValue || String(e.oldValue).trim() === "")) {
      sendWelcomeOfferLetter(sheet, row);
    } else {
      var fieldNames = {
        6: "Teacher Name",
        7: "Gender",
        8: "Base Salary",
        9: "Department",
        10: "Joining Date",
        11: "Residential Address",
        12: "Bonus Amount",
        13: "Mobile Phone Number",
        14: "Registered Email Address"
      };

      var fieldName = fieldNames[col] || "Faculty Record";
      var oldValDisplay = (e.oldValue !== undefined && e.oldValue !== null) ? String(e.oldValue).trim() : "Previous Value";
      var newValDisplay = (e.value !== undefined && e.value !== null) ? String(e.value).trim() : String(e.range.getValue()).trim();

      if (col === COL_SALARY || col === COL_BONUS) {
        if (!isNaN(Number(oldValDisplay))) oldValDisplay = "₹" + Number(oldValDisplay).toLocaleString('en-IN');
        if (!isNaN(Number(newValDisplay))) newValDisplay = "₹" + Number(newValDisplay).toLocaleString('en-IN');
      }

      sendStatementForRow(sheet, row, fieldName, oldValDisplay, newValDisplay);
    }
  }
}

function processContactUsStatusChange(row, statusVal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cuSheet = ss.getSheetByName(SHEET_CONTACT);
  if (!cuSheet) return;

  var normStatus = String(statusVal || "").trim().toUpperCase();
  if (normStatus === "RESOLVED" || normStatus === "REJECTED") {
    var reqData = cuSheet.getRange(row, 1, 1, 10).getValues()[0];
    var reqId = reqData[0];
    var facultyName = reqData[3];
    var email = String(reqData[5] || "").trim();
    var reqType = reqData[6];
    var reqDesc = reqData[7];

    if (email) {
      var statusColor = normStatus === "RESOLVED" ? "#16a34a" : "#dc2626";
      var statusTitle = normStatus === "RESOLVED" ? "Resolved & Completed" : "Declined / Rejected";

      var statusEmailHtml = "<div style='font-family:Arial,sans-serif; padding:20px; border:1px solid #e2e8f0; border-radius:10px; max-width:600px;'>" +
        "<h2 style='color:" + statusColor + "; margin-top:0;'>Request " + reqId + " " + statusTitle + "</h2>" +
        "<p>Dear <strong>" + facultyName + "</strong>,</p>" +
        "<p>Your faculty assistance ticket has been reviewed by the administration and marked as <strong>" + normStatus + "</strong>.</p>" +
        "<table style='width:100%; border-collapse:collapse; margin:15px 0; font-size:13px;'>" +
        "<tr style='background:#f8fafc;'><td style='padding:8px; border:1px solid #e2e8f0;'>Request ID</td><td style='padding:8px; border:1px solid #e2e8f0;'><strong>" + reqId + "</strong></td></tr>" +
        "<tr><td style='padding:8px; border:1px solid #e2e8f0;'>Request Category</td><td style='padding:8px; border:1px solid #e2e8f0;'>" + reqType + "</td></tr>" +
        "<tr style='background:#f8fafc;'><td style='padding:8px; border:1px solid #e2e8f0;'>Original Query</td><td style='padding:8px; border:1px solid #e2e8f0;'>" + reqDesc + "</td></tr>" +
        "<tr><td style='padding:8px; border:1px solid #e2e8f0;'>Final Decision</td><td style='padding:8px; border:1px solid #e2e8f0;'><strong style='color:" + statusColor + "; font-size:14px;'>" + normStatus + "</strong></td></tr>" +
        "</table>" +
        "<p style='font-size:12.5px; color:#64748b;'>In accordance with service record archival policy, this ticket will be cleared from active records in 30 seconds.</p>" +
        "<div style='background:#fef2f2; border:1px solid #fecaca; padding:10px; border-radius:6px; font-size:11px; color:#991b1b; text-align:center; font-weight:700; margin-top:20px;'>" +
        "⚠️ THIS IS A SYSTEM GENERATED NOTIFICATION EMAIL. PLEASE DO NOT REPLY. IF RECEIVED BY MISTAKE, PLEASE DESTROY THIS EMAIL IMMEDIATELY." +
        "</div></div>";

      MailApp.sendEmail({
        to: email,
        subject: "Request " + reqId + " Status Update: " + normStatus,
        htmlBody: statusEmailHtml
      });
    }

    ss.toast("Request " + reqId + " marked " + normStatus + ". Deleting in 30s...", "⏳ 30s Countdown", 6);
    Utilities.sleep(30000);
    SpreadsheetApp.flush();

    var currentStatus = String(cuSheet.getRange(row, CU_STATUS).getValue() || "").trim().toUpperCase();
    if (currentStatus === "RESOLVED" || currentStatus === "REJECTED") {
      cuSheet.deleteRow(row);
      ss.toast("Request " + reqId + " permanently deleted from CONTACT US.", "✅ Ticket Purged", 4);
    }
  }
}

function processDeletionAndMoveToBin(row, statusVal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var binSheet = ss.getSheetByName(SHEET_BIN);
  if (!mainSheet || !binSheet) return;

  var rowData = mainSheet.getRange(row, COL_UID, 1, 11).getValues()[0];
  var uid = rowData[0];
  var teacherName = rowData[1];
  var deptVal = rowData[4] || "Academic Faculty";
  var dateVal = rowData[5];
  var addressVal = String(rowData[6] || "N/A");
  var salaryVal = Number(rowData[3]) || 0;
  var bonusVal = Number(rowData[7]) || 0;
  var phoneVal = String(rowData[8] || "");
  var email = String(rowData[9] || "").trim();

  if (!uid) return;

  var formattedDate = dateVal instanceof Date 
    ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd MMMM yyyy") 
    : (dateVal || "N/A");
  var currentTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm:ss a");
  var currentDateOnly = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy");
  var randomIp = "192.168." + Math.floor(Math.random() * 200 + 10) + "." + Math.floor(Math.random() * 250 + 2);
  var sixDigitRef = Math.floor(100000 + Math.random() * 900000).toString();

  var cleanPhone = phoneVal.replace(/\D/g, "");
  var maskedPhone = cleanPhone.length > 4 ? "***************" + cleanPhone.slice(-4) : phoneVal;
  var maskedEmail = email.length > 15 ? "***************" + email.slice(-15) : email;

  if (email) {
    var exitPdfBlob = buildTwoPagePdfBlob(
      teacherName, 
      uid, 
      deptVal, 
      formattedDate, 
      addressVal,
      salaryVal, 
      bonusVal, 
      (salaryVal + bonusVal), 
      maskedPhone, 
      maskedEmail, 
      currentDateOnly, 
      currentTimestamp, 
      randomIp, 
      sixDigitRef, 
      "Official Faculty Relieving & Record Closure Statement", 
      statusVal
    );

    var exitNoticeHtml = "<!DOCTYPE html><html><body style='margin:0; padding:20px; font-family:Arial,sans-serif; background-color:#f8fafc;'>" +
      "<div style='max-width:600px; margin:auto; background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:24px; box-shadow:0 4px 15px rgba(0,0,0,0.05);'>" +
      "<h3 style='color:#b91c1c; margin-top:0; font-size:18px;'>Faculty Record Closure Notice</h3>" +
      "<p style='color:#334155; font-size:14px; line-height:1.5;'>Dear <strong>" + teacherName + "</strong>,</p>" +
      "<p style='color:#334155; font-size:14px; line-height:1.5;'>Your profile has been marked as <strong>" + statusVal + "</strong>. Attached is your official <strong>2-Page Faculty Relieving & Record Closure Statement (PDF)</strong>.</p>" +
      "<p style='color:#334155; font-size:14px; line-height:1.5;'>This record has been transferred to <strong>BIN RECORDS</strong>. It will be permanently purged in 30 seconds unless marked as <strong>Restore</strong>.</p>" +
      "<br><div style='background:#fef2f2; border:1px solid #fecaca; padding:12px; border-radius:8px; font-size:11px; color:#991b1b; text-align:center; font-weight:700; line-height:1.4;'>" +
      "⚠️ THIS IS A SYSTEM GENERATED NOTIFICATION EMAIL. PLEASE DO NOT REPLY. IF RECEIVED BY MISTAKE, PLEASE DESTROY THIS EMAIL IMMEDIATELY." +
      "</div></div></body></html>";

    MailApp.sendEmail({
      to: email,
      subject: "Official Relieving & Exit Notice: " + teacherName + " (UID: " + uid + ")",
      htmlBody: exitNoticeHtml,
      attachments: [exitPdfBlob]
    });
  }

  var binData = rowData.slice();
  binData[10] = "Deleted";

  var nextBinRow = getNextAvailableRow(binSheet, COL_UID, HEADER_ROW + 1);
  binSheet.getRange(nextBinRow, COL_UID, 1, 11).setValues([binData]);
  binSheet.getRange(nextBinRow, COL_BIN_TIME).setValue(new Date());

  mainSheet.deleteRow(row);
  ss.toast("Record moved to BIN RECORDS. Will purge in 30s unless restored.", "🗑️ Recycled", 5);

  Utilities.sleep(30000);
  SpreadsheetApp.flush();
  permanentlyPurgeIfUnrestored(uid);
}

function restoreRowToMainSheet(binRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var binSheet = ss.getSheetByName(SHEET_BIN);
  if (!mainSheet || !binSheet) return;

  var rowData = binSheet.getRange(binRow, COL_UID, 1, 11).getValues()[0];
  if (!rowData[0]) return;

  rowData[10] = "Active";
  var nextRow = getNextAvailableRow(mainSheet, COL_UID, HEADER_ROW + 1);
  mainSheet.getRange(nextRow, COL_UID, 1, 11).setValues([rowData]);
  binSheet.deleteRow(binRow);
  ss.toast("Record " + rowData[0] + " restored to Sheet1!", "🛡️ Restored", 4);
}

function permanentlyPurgeIfUnrestored(targetUid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var binSheet = ss.getSheetByName(SHEET_BIN);
  if (!binSheet || binSheet.getLastRow() <= HEADER_ROW) return;

  var data = binSheet.getRange(HEADER_ROW + 1, COL_UID, binSheet.getLastRow() - HEADER_ROW, 11).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === String(targetUid).trim()) {
      if (String(data[i][10]).toUpperCase() !== "RESTORE") {
        binSheet.deleteRow(HEADER_ROW + 1 + i);
        ss.toast("Record " + targetUid + " permanently purged from Bin.", "🗑️ Purged", 4);
      }
      break;
    }
  }
}

function purgeExpiredBinRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var binSheet = ss.getSheetByName(SHEET_BIN);
  if (!binSheet || binSheet.getLastRow() <= HEADER_ROW) return;

  var data = binSheet.getRange(HEADER_ROW + 1, COL_UID, binSheet.getLastRow() - HEADER_ROW, 12).getValues();
  var now = new Date().getTime();
  var count = 0;

  for (var i = data.length - 1; i >= 0; i--) {
    var status = String(data[i][10] || "").toUpperCase();
    var time = data[i][11];
    if (status !== "RESTORE" && time instanceof Date) {
      if ((now - time.getTime()) / 1000 >= 30) {
        binSheet.deleteRow(HEADER_ROW + 1 + i);
        count++;
      }
    }
  }
  ss.toast("Purged " + count + " expired bin records.", "🗑️ Purge Complete", 4);
}

function calculateBonusFromJoiningDate(dateInput) {
  var joinDate = dateInput instanceof Date ? dateInput : new Date(dateInput);
  if (isNaN(joinDate.getTime())) return 0;

  var today = new Date();
  var monthsDiff = (today.getFullYear() - joinDate.getFullYear()) * 12 + (today.getMonth() - joinDate.getMonth());
  if (today.getDate() < joinDate.getDate()) monthsDiff--;
  if (monthsDiff < 3) return 0;

  var yearsDiff = Math.floor(monthsDiff / 12);
  return Math.min(150000, 15000 + (yearsDiff * 15000));
}

function generateUnique6DigitUid(sheet, colIndex) {
  var firstDataRow = HEADER_ROW + 1;
  var lastRow = sheet.getLastRow();
  var values = lastRow >= firstDataRow
    ? sheet.getRange(firstDataRow, colIndex, lastRow - HEADER_ROW, 1).getValues().flat().map(String)
    : [];
  var set = new Set(values);
  var newUid;
  do {
    newUid = Math.floor(100000 + Math.random() * 900000).toString();
  } while (set.has(newUid));
  return newUid;
}

function getNextAvailableRow(sheet, checkCol, startRow) {
  var max = sheet.getMaxRows();
  if (max < startRow) sheet.insertRowsAfter(max, 50);
  var colVals = sheet.getRange(startRow, checkCol, sheet.getMaxRows() - startRow + 1, 1).getValues();
  for (var i = 0; i < colVals.length; i++) {
    if (!colVals[i][0]) return startRow + i;
  }
  var firstInsertedRow = sheet.getMaxRows() + 1;
  sheet.insertRowsAfter(sheet.getMaxRows(), 10);
  return firstInsertedRow;
}

function buildTwoPagePdfBlob(name, uid, dept, joiningDate, address, salary, bonus, totalSalary, maskedPhone, maskedEmail, currentDateOnly, currentTimestamp, randomIp, sixDigitRef, docTitle, statusText) {
  var latexHtml = "<!DOCTYPE html><html><head><meta charset='utf-8'>" +
    "<style>@page { size: A4; margin: 18mm; } body { font-family: 'Times New Roman', Times, serif; font-size: 10.5pt; line-height: 1.5; color: #111; margin:0; padding:0; }" +
    ".page-break { page-break-before: always; } .inst-header { text-align: center; margin-bottom: 12px; }" +
    ".inst-title { font-size: 17.5pt; font-weight: bold; letter-spacing: 1px; text-transform: uppercase; margin: 0; }" +
    ".thick-rule { border: none; border-top: 1.8pt solid #111; margin: 6px 0 10px 0; }" +
    ".doc-title { font-size: 12.5pt; font-weight: bold; text-transform: uppercase; margin: 6px 0; }" +
    ".section-heading { font-size: 11pt; font-weight: bold; text-transform: uppercase; border-bottom: 1pt solid #111; padding-bottom: 3px; margin: 16px 0 8px 0; }" +
    ".grid-table { width: 100%; border-collapse: collapse; margin: 8px 0 14px 0; font-size: 9.5pt; }" +
    ".grid-table th { background-color: #f8fafc; border: 1px solid #111; padding: 5px 8px; text-align: left; }" +
    ".grid-table td { border: 1px solid #111; padding: 5px 8px; }" +
    ".total-row { background-color: #f1f5f9; font-weight: bold; }" +
    ".sig-table { margin-top: 25px; width: 100%; border-collapse: collapse; }" +
    ".sig-table td { vertical-align: top; width: 50%; padding-right: 15px; }" +
    ".times-italic-sig { font-family: 'Times New Roman', Times, serif; font-style: italic; font-size: 14pt; color: #0b192c; margin: 4px 0 2px 0; }" +
    ".sig-line { border-bottom: 1pt solid #111; width: 85%; margin-bottom: 4px; }" +
    ".green-seal { display: inline-block; border: 1.5px dashed #166534; background-color: #f0fdf4; padding: 4px 10px; border-radius: 4px; margin-top: 4px; color: #166534; font-size: 7.5pt; font-weight: bold; text-transform: uppercase; }" +
    ".blue-seal { display: inline-block; border: 1.5px dashed #1e40af; background-color: #eff6ff; padding: 4px 10px; border-radius: 4px; margin-top: 4px; color: #1e40af; font-size: 7.5pt; font-weight: bold; text-transform: uppercase; }" +
    ".audit-box { margin-top: 45px; border-top: 1pt solid #94a3b8; padding: 12px; background-color: #f8fafc; width: 100%; border-collapse: collapse; text-align: center; font-size: 8.5pt; color: #475569; }" +
    "</style></head><body>" +
    "<div class='inst-header'><div class='inst-title'>Institute of Academic Excellence</div><hr class='thick-rule' /><div class='doc-title'>" + (docTitle || "Faculty Statement") + "</div></div>" +
    "<p><strong>Date:</strong> " + currentDateOnly + " | <strong>Candidate:</strong> " + name + " (UID: <code>" + uid + "</code>) | <strong>Dept:</strong> " + dept + "</p>" +
    "<div class='section-heading'>1. Appointment Highlights & Verified Credentials</div>" +
    "<table class='grid-table'><thead><tr><th>Parameter</th><th>Registry Entry</th></tr></thead><tbody>" +
    "<tr><td>Name</td><td><strong>" + name + "</strong></td></tr><tr><td>UID</td><td><code>" + uid + "</code></td></tr>" +
    "<tr><td>Department</td><td>" + dept + "</td></tr><tr><td>Joining Date</td><td>" + joiningDate + "</td></tr>" +
    "<tr><td>Residential Address</td><td>" + (address || "N/A") + "</td></tr>" +
    "<tr><td>Mobile (Secured)</td><td><code>" + maskedPhone + "</code></td></tr>" +
    "<tr><td>Email (Secured - Last 15)</td><td><code>" + maskedEmail + "</code></td></tr>" +
    "<tr><td>Status</td><td><strong style='color:#166534;'>" + (statusText || "ACTIVE") + "</strong></td></tr></tbody></table>" +
    "<div class='section-heading'>2. Remuneration Schedule</div>" +
    "<table class='grid-table'><thead><tr><th>Component</th><th style='text-align:right;'>Monthly Sum (INR)</th></tr></thead><tbody>" +
    "<tr><td>Base Salary</td><td style='text-align:right;'>₹" + salary.toLocaleString('en-IN') + "</td></tr>" +
    "<tr><td>Experience Bonus</td><td style='text-align:right; color:#166534; font-weight:bold;'>₹" + bonus.toLocaleString('en-IN') + "</td></tr>" +
    "<tr class='total-row'><td>TOTAL MONTHLY GROSS</td><td style='text-align:right;'>₹" + totalSalary.toLocaleString('en-IN') + "</td></tr></tbody></table>" +
    "<div class='page-break'></div>" +
    "<div class='inst-header'><div class='doc-title'>Terms of Governance & Institutional Endorsement</div><hr class='thick-rule' /></div>" +
    "<div class='section-heading'>3. Terms & Conditions</div>" +
    "<p style='font-size:9.8pt; line-height:1.5;'>This statement reflects verified institutional records. Bonus milestones are governed by academic experience guidelines.</p>" +
    "<div class='section-heading'>4. Institutional Signatures</div>" +
    "<table class='sig-table'><tr>" +
    "<td><div class='times-italic-sig'>J. Upadhyay</div><div class='sig-line'></div><strong>Jagneswar Upadhyay</strong><br><span style='font-size:8.5pt;'>Dean of Academic Affairs</span><br><div class='green-seal'>✔ AUTO SEAL VERIFIED IN GREEN TEX</div></td>" +
    "<td><div class='times-italic-sig'>" + name + "</div><div class='sig-line'></div><strong>Accepted & Acknowledged</strong><br><span style='font-size:8.5pt;'>Candidate Digital Signature</span></td>" +
    "</tr></table>" +
    "<div class='audit-box'>IP: <strong>" + randomIp + "</strong> | Ref: <strong>REF-" + sixDigitRef + "</strong> | Certified: <strong>" + currentTimestamp + "</strong></div>" +
    "</body></html>";

  return Utilities.newBlob(latexHtml, "text/html", "Statement_" + uid + ".html")
    .getAs("application/pdf")
    .setName("Faculty_Statement_2Page_" + uid + ".pdf");
}

/**
 * Builds responsive email template showing BEFORE and AFTER modification tables.
 * Email displays only the last 15 characters masked.
 */
function buildResponsiveEmailHtml(badge, title, name, uid, dept, date, address, sal, bon, tot, phone, email, text, ref, status, log) {
  var logHtml = "";
  if (log && log.fieldName) {
    logHtml = "<div style='background:#f0fdf4; border:2px solid #22c55e; border-radius:10px; padding:16px; margin:18px 0; box-shadow:0 2px 8px rgba(0,0,0,0.03);'>" +
      "<div style='font-size:11.5px; font-weight:800; color:#15803d; text-transform:uppercase; letter-spacing:0.8px;'>🔔 OFFICIAL FIELD MODIFICATION LOG</div>" +
      "<p style='margin:8px 0 10px 0; font-size:14px; color:#1e293b; line-height:1.5;'>" +
      "Your <strong>" + log.fieldName + "</strong> has been updated in our official records as follows:" +
      "</p>" +
      "<table width='100%' style='border-collapse:collapse; margin-top:10px; font-size:13px;'>" +
      "<tr>" +
      "<td width='50%' style='padding:12px 14px; background:#fee2e2; border:1px solid #fecaca; border-radius:6px; color:#991b1b; vertical-align:top;'>" +
      "<span style='font-size:10.5px; font-weight:800; text-transform:uppercase; letter-spacing:0.5px; display:block; margin-bottom:4px;'>🔴 PREVIOUS (OLD) VALUE:</span>" +
      "<span style='font-size:14.5px; font-weight:700; font-family:monospace; word-break:break-all;'>" + log.oldVal + "</span>" +
      "</td>" +
      "<td width='4%'>&nbsp;</td>" +
      "<td width='46%' style='padding:12px 14px; background:#dcfce7; border:1px solid #bbf7d0; border-radius:6px; color:#166534; vertical-align:top;'>" +
      "<span style='font-size:10.5px; font-weight:800; text-transform:uppercase; letter-spacing:0.5px; display:block; margin-bottom:4px;'>🟢 UPDATED (NEW) VALUE:</span>" +
      "<span style='font-size:14.5px; font-weight:700; font-family:monospace; word-break:break-all;'>" + log.newVal + "</span>" +
      "</td>" +
      "</tr>" +
      "</table>" +
      "</div>";
  }

  return "<!DOCTYPE html><html><body style='margin:0; padding:15px; background:#f8fafc; font-family:Arial,sans-serif;'>" +
    "<div style='max-width:600px; margin:auto; background:#fff; border:1px solid #e2e8f0; border-radius:12px; overflow:hidden;'>" +
    "<div style='background:#1e1b4b; padding:25px; color:#fff;'><span style='background:rgba(255,255,255,0.2); padding:3px 8px; border-radius:10px; font-size:11px; font-weight:bold;'>" + badge + "</span>" +
    "<h2 style='margin:10px 0 0 0; font-size:20px;'>" + title + "</h2><p style='margin:4px 0 0 0; font-size:12px; color:#c7d2fe;'>UID: " + uid + " | Ref: REF-" + ref + "</p></div>" +
    "<div style='padding:22px; color:#334155; font-size:13.5px; line-height:1.5;'>" +
    "<p>Dear <strong>" + name + "</strong>,</p><p>" + text + "</p>" +
    logHtml +
    "<table width='100%' style='border-collapse:collapse; margin:15px 0; font-size:13px;'>" +
    "<tr style='background:#f1f5f9;'><td style='padding:8px; border:1px solid #e2e8f0;'>Department</td><td style='padding:8px; border:1px solid #e2e8f0;'><strong>" + dept + "</strong></td></tr>" +
    "<tr><td style='padding:8px; border:1px solid #e2e8f0;'>Residential Address</td><td style='padding:8px; border:1px solid #e2e8f0;'>" + (address || "N/A") + "</td></tr>" +
    "<tr style='background:#f1f5f9;'><td style='padding:8px; border:1px solid #e2e8f0;'>Base Salary</td><td style='padding:8px; border:1px solid #e2e8f0;'>₹" + sal.toLocaleString('en-IN') + "</td></tr>" +
    "<tr style='background:#f0fdf4;'><td style='padding:8px; border:1px solid #e2e8f0; color:#15803d;'>Experience Bonus</td><td style='padding:8px; border:1px solid #e2e8f0; color:#15803d;'><strong>₹" + bon.toLocaleString('en-IN') + "</strong></td></tr>" +
    "<tr style='background:#4338ca; color:#fff;'><td style='padding:10px; font-weight:bold;'>TOTAL DISBURSEMENT</td><td style='padding:10px; font-weight:bold; font-size:15px;'>₹" + tot.toLocaleString('en-IN') + "</td></tr></table>" +
    "<div style='background:#fef2f2; border:1px solid #fecaca; padding:10px; border-radius:6px; font-size:11px; color:#991b1b; text-align:center; font-weight:700; margin-top:20px;'>" +
    "⚠️ THIS IS A SYSTEM GENERATED NOTIFICATION EMAIL. PLEASE DO NOT REPLY. IF RECEIVED BY MISTAKE, PLEASE DESTROY THIS EMAIL IMMEDIATELY.</div>" +
    "</div></div></body></html>";
}

function sendWelcomeOfferLetter(sheet, row) {
  var d = sheet.getRange(row, COL_UID, 1, 11).getValues()[0];
  if (!d[9]) return;
  var sal = Number(d[3]) || 0; 
  var address = String(d[6] || "N/A");
  var bon = Number(d[7]) || 0;
  var dateStr = d[5] instanceof Date ? Utilities.formatDate(d[5], Session.getScriptTimeZone(), "dd/MM/yyyy") : d[5];
  var phone = String(d[8] || ""); 
  var email = String(d[9] || "");
  var maskedP = phone.length > 4 ? "***************" + phone.slice(-4) : phone;
  var maskedE = email.length > 15 ? "***************" + email.slice(-15) : email;
  var ref = Math.floor(100000 + Math.random() * 900000).toString();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a");

  var pdf = buildTwoPagePdfBlob(d[1], d[0], d[4], dateStr, address, sal, bon, (sal+bon), maskedP, maskedE, now, now, "192.168.1.1", ref, "Formal Appointment Letter", "ACTIVE");
  var html = buildResponsiveEmailHtml("Appointment", "Welcome to the Faculty", d[1], d[0], d[4], dateStr, address, sal, bon, (sal+bon), maskedP, maskedE, "Your faculty profile has been created. Attached is your official 2-Page Verified Appointment Statement (PDF).", ref, "ACTIVE", null);

  MailApp.sendEmail({ to: email, subject: "Welcome to Faculty: " + d[1] + " (" + d[0] + ")", htmlBody: html, attachments: [pdf] });
}

function sendStatementForRow(sheet, row, fName, oldV, newV) {
  var d = sheet.getRange(row, COL_UID, 1, 11).getValues()[0];
  if (!d[9]) return;
  var sal = Number(d[3]) || 0; 
  var address = String(d[6] || "N/A");
  var bon = Number(d[7]) || 0;
  var dateStr = d[5] instanceof Date ? Utilities.formatDate(d[5], Session.getScriptTimeZone(), "dd/MM/yyyy") : d[5];
  var phone = String(d[8] || ""); 
  var email = String(d[9] || "").trim();
  var maskedP = phone.length > 4 ? "***************" + phone.slice(-4) : phone;
  var maskedE = email.length > 15 ? "***************" + email.slice(-15) : email;
  var ref = Math.floor(100000 + Math.random() * 900000).toString();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a");

  var pdf = buildTwoPagePdfBlob(d[1], d[0], d[4], dateStr, address, sal, bon, (sal+bon), maskedP, maskedE, now, now, "192.168.1.1", ref, "Updated Compensation Statement", d[10]);
  var log = fName ? { fieldName: fName, oldVal: oldV, newVal: newV } : null;
  var html = buildResponsiveEmailHtml("Record Updated", "Faculty Record Statement", d[1], d[0], d[4], dateStr, address, sal, bon, (sal+bon), maskedP, maskedE, "Your faculty details were updated in the registry.", ref, d[10], log);

  MailApp.sendEmail({ to: email, subject: "Record Updated (" + (fName || "Statement") + "): " + d[1], htmlBody: html, attachments: [pdf] });

  if (fName === "Registered Email Address" && oldV && oldV.includes("@") && oldV !== email) {
    try {
      MailApp.sendEmail({ to: oldV, subject: "Security Notice: Registered Email Address Changed for " + d[1], htmlBody: html, attachments: [pdf] });
    } catch(err){}
  }
}

function sendStatementsForExistingRecords() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  for (var r = 3; r <= s.getLastRow(); r++) {
    if (s.getRange(r, COL_STATUS).getValue() === "Active" && s.getRange(r, COL_EMAIL).getValue()) {
      sendStatementForRow(s, r, "Bulk Statement Dispatch", "Current", "Verified");
    }
  }
}

function sendWelcomeOfferForSelectedRow() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var r = s.getActiveCell().getRow();
  if (r > HEADER_ROW) sendWelcomeOfferLetter(s, r);
}

function sendTestEmailNow() {
  var myEmail = Session.getActiveUser().getEmail() || "upadhyayankur303@gmail.com";
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a");
  var pdf = buildTwoPagePdfBlob("Prof. Test Faculty", "749201", "Computer Science", "30/06/2017", "LPU Campus, Phagwara, Punjab", 70000, 15000, 85000, "***************8411", "***************" + myEmail.slice(-15), now, now, "192.168.1.108", "849201", "Appointment Letter", "ACTIVE");
  var testLog = { fieldName: "Mobile Phone Number", oldVal: "9434868411", newVal: "9876543210" };
  var html = buildResponsiveEmailHtml("Test Record Update", "Faculty Verification Notice", "Prof. Test Faculty", "749201", "Computer Science", "30/06/2017", "LPU Campus, Phagwara, Punjab", 70000, 15000, 85000, "***************8411", "***************" + myEmail.slice(-15), "Testing the before-and-after audit log rendering.", "849201", "ACTIVE", testLog);
  MailApp.sendEmail({ to: myEmail, subject: "✅ Test Verification (Old vs New Value Preview)", htmlBody: html, attachments: [pdf] });
  SpreadsheetApp.getActiveSpreadsheet().toast("Test email sent to " + myEmail, "Dispatched", 4);
}

/**
 * --------------------------------------------------------------------------
 * EMBEDDED WEB PORTAL HTML (SUPPORTS ADDRESS & 10-CHAR EMAIL MASKING)
 * --------------------------------------------------------------------------
 */
function getPortalHtml() {
  return '<!DOCTYPE html>\n' +
'<html lang="en">\n' +
'<head>\n' +
'  <meta charset="UTF-8">\n' +
'  <title>Faculty Management Portal - Lovely Professional University</title>\n' +
'  <meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
'  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">\n' +
'  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>\n' +
'  <style>\n' +
'    * { box-sizing: border-box; margin: 0; padding: 0; font-family: "Inter", sans-serif; }\n' +
'    body { background-color: #f1f5f9; color: #1e293b; min-height: 100vh; display: flex; flex-direction: column; }\n' +
'    .browser-banner { background-color: #1e1b4b; color: #c7d2fe; font-size: 11.5px; padding: 6px 16px; text-align: center; font-weight: 500; }\n' +
'    .browser-banner a { color: #ffffff; text-decoration: underline; font-weight: 700; }\n' +
'    header { background: linear-gradient(135deg, #1e1b4b 0%, #312e81 60%, #4338ca 100%); color: #fff; padding: 16px 24px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 14px rgba(0,0,0,0.12); }\n' +
'    header h1 { font-size: 19px; font-weight: 800; }\n' +
'    .header-right { display: flex; align-items: center; gap: 16px; }\n' +
'    .live-clock { font-size: 12.5px; color: #c7d2fe; font-weight: 600; background: rgba(255,255,255,0.1); padding: 5px 12px; border-radius: 20px; }\n' +
'    .btn-logout { background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.3); color: #fff; padding: 6px 14px; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 600; }\n' +
'    .container { max-width: 1100px; width: 100%; margin: 25px auto; padding: 0 16px; flex: 1; }\n' +
'    .card { background: #fff; border-radius: 14px; border: 1px solid #e2e8f0; box-shadow: 0 4px 20px rgba(0,0,0,0.05); padding: 30px; }\n' +
'    .login-box { max-width: 440px; margin: 35px auto; }\n' +
'    .form-group { margin-bottom: 16px; }\n' +
'    label { display: block; font-size: 13px; font-weight: 600; color: #475569; margin-bottom: 6px; }\n' +
'    select, input, textarea { width: 100%; padding: 10px 14px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 14px; outline: none; }\n' +
'    select:focus, input:focus, textarea:focus { border-color: #4338ca; box-shadow: 0 0 0 3px rgba(67, 56, 202, 0.15); }\n' +
'    .btn-primary { width: 100%; background: #4338ca; color: #fff; border: none; padding: 12px; border-radius: 8px; font-size: 14px; font-weight: 700; cursor: pointer; margin-top: 8px; }\n' +
'    .btn-primary:hover { background: #3730a3; }\n' +
'    .timer-row { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; }\n' +
'    .timer-text { font-size: 13px; font-weight: 700; color: #dc2626; }\n' +
'    .btn-resend { background: #e2e8f0; color: #94a3b8; border: 1px solid #cbd5e1; padding: 6px 14px; border-radius: 6px; font-size: 12px; font-weight: 700; cursor: not-allowed; }\n' +
'    .btn-resend.active { background: #e0e7ff; color: #4338ca; border-color: #c7d2fe; cursor: pointer; }\n' +
'    .alert-box { padding: 12px 14px; border-radius: 8px; font-size: 13px; margin-top: 14px; display: none; line-height: 1.45; }\n' +
'    .alert-error { background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; }\n' +
'    .alert-success { background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }\n' +
'    .nav-tabs { display: flex; gap: 8px; margin-bottom: 20px; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; flex-wrap: wrap; }\n' +
'    .tab-btn { background: none; border: none; font-size: 14px; font-weight: 600; color: #64748b; padding: 8px 16px; cursor: pointer; border-radius: 6px; }\n' +
'    .tab-btn.active { background: #e0e7ff; color: #4338ca; }\n' +
'    .badge { background: #dc2626; color: #fff; padding: 2px 7px; border-radius: 12px; font-size: 11px; margin-left: 6px; font-weight: 800; }\n' +
'    .table-responsive { overflow-x: auto; margin-top: 15px; }\n' +
'    table { width: 100%; border-collapse: collapse; font-size: 13px; text-align: left; }\n' +
'    th { background: #f8fafc; padding: 10px 12px; border-bottom: 2px solid #e2e8f0; color: #475569; font-weight: 700; }\n' +
'    td { padding: 10px 12px; border-bottom: 1px solid #f1f5f9; color: #1e293b; }\n' +
'    .btn-sm { padding: 5px 10px; font-size: 11.5px; border-radius: 4px; border: none; cursor: pointer; font-weight: 600; margin-right: 4px; }\n' +
'    .btn-edit { background: #e0e7ff; color: #3730a3; }\n' +
'    .btn-bin { background: #fee2e2; color: #991b1b; }\n' +
'    .btn-restore { background: #dcfce7; color: #166534; }\n' +
'    .btn-resolve { background: #dcfce7; color: #166534; }\n' +
'    .btn-reject { background: #fee2e2; color: #991b1b; }\n' +
'    .grid-form { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 12px; }\n' +
'    .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 14px; margin: 20px 0; }\n' +
'    .kpi-box { background: #fff; border: 1px solid #e2e8f0; padding: 16px; border-radius: 10px; }\n' +
'    .kpi-title { font-size: 11px; text-transform: uppercase; color: #64748b; font-weight: 700; }\n' +
'    .kpi-val { font-size: 21px; font-weight: 800; color: #1e293b; margin-top: 4px; }\n' +
'    .profile-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 20px; margin-top: 15px; }\n' +
'    .upload-zone { border: 2px dashed #cbd5e1; border-radius: 10px; padding: 30px; text-align: center; background: #f8fafc; cursor: pointer; }\n' +
'    .upload-zone:hover { border-color: #4338ca; background: #f1f5f9; }\n' +
'    footer { text-align: center; padding: 22px; font-size: 12px; color: #64748b; }\n' +
'    @media only screen and (max-width: 600px) {\n' +
'      .container { padding: 0 8px; margin: 15px auto; }\n' +
'      .card { padding: 20px 16px; }\n' +
'      .grid-form, .kpi-row { grid-template-columns: 1fr; }\n' +
'    }\n' +
'  </style>\n' +
'</head>\n' +
'<body>\n' +
'  <div class="browser-banner">💡 For optimal viewing, download the latest version of <a href="https://www.google.com/chrome/" target="_blank">Google Chrome</a> or <a href="https://www.mozilla.org/firefox/" target="_blank">Mozilla Firefox</a>.</div>\n' +
'  <header>\n' +
'    <h1>🏫 Lovely Professional University Portal</h1>\n' +
'    <div class="header-right">\n' +
'      <span id="liveClockDisplay" class="live-clock">🕒 Time</span>\n' +
'      <button id="logoutBtn" class="btn-logout" style="display:none;" onclick="logout()">Sign Out</button>\n' +
'    </div>\n' +
'  </header>\n' +
'  <div class="container">\n' +
'    <div id="loginSection" class="card login-box">\n' +
'      <h2 style="font-size:20px; font-weight:800; margin-bottom:18px; text-align:center;">Secure Portal Login</h2>\n' +
'      <div class="form-group">\n' +
'        <label>Select Login Role</label>\n' +
'        <select id="loginRole" onchange="switchRoleFields()">\n' +
'          <option value="STAFF">Faculty / Staff Login</option>\n' +
'          <option value="ADMIN">Administrator Login</option>\n' +
'        </select>\n' +
'      </div>\n' +
'      <div id="adminFields" style="display:none;">\n' +
'        <div class="form-group"><label>Admin User ID</label><input type="text" id="adminUser" placeholder="Enter User ID (001)"></div>\n' +
'        <div class="form-group"><label>Admin Password</label><input type="password" id="adminPass" placeholder="Enter Password (001)"></div>\n' +
'      </div>\n' +
'      <div id="staffFields">\n' +
'        <div class="form-group"><label>Registered Mobile or UID</label><input type="text" id="staffIdentifier" placeholder="Enter UID or Mobile"></div>\n' +
'        <div class="form-group"><label>Date of Joining (Password)</label><input type="date" id="staffJoiningDate"></div>\n' +
'      </div>\n' +
'      <div id="otpStep" style="display:none;">\n' +
'        <div class="form-group">\n' +
'          <label>Enter 6-Digit OTP</label>\n' +
'          <input type="text" id="staffOtpInput" placeholder="Enter OTP" maxlength="6" style="letter-spacing:5px; font-weight:800; text-align:center; font-size:20px;">\n' +
'        </div>\n' +
'        <div class="timer-row">\n' +
'          <span id="timerDisplay" class="timer-text">⏱️ Expires in: 02:00</span>\n' +
'          <button id="resendOtpBtn" type="button" class="btn-resend" disabled onclick="handleResendOtp()">Resend OTP</button>\n' +
'        </div>\n' +
'      </div>\n' +
'      <button id="loginSubmitBtn" class="btn-primary" onclick="handleLoginSubmit()">Continue</button>\n' +
'      <div id="loginAlert" class="alert-box"></div>\n' +
'    </div>\n' +
'    <div id="adminSection" class="card" style="display:none;">\n' +
'      <div class="nav-tabs">\n' +
'        <button class="tab-btn active" onclick="switchAdminTab(\'ACTIVE_TAB\', this)">Active Faculty</button>\n' +
'        <button class="tab-btn" onclick="switchAdminTab(\'ADD_TAB\', this)">+ Add Single Faculty</button>\n' +
'        <button class="tab-btn" onclick="switchAdminTab(\'BULK_TAB\', this)">📁 Bulk Upload (Excel/CSV)</button>\n' +
'        <button class="tab-btn" onclick="switchAdminTab(\'BIN_TAB\', this)">🗑️ Bin Records</button>\n' +
'        <button class="tab-btn" onclick="switchAdminTab(\'REQUESTS_TAB\', this)">📬 Support Tickets <span id="pendingBadge" class="badge">0</span></button>\n' +
'      </div>\n' +
'      <div id="adminActiveTab">\n' +
'        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">\n' +
'          <h3>Active Faculty Registry (Sheet1)</h3>\n' +
'          <button class="btn-sm btn-edit" onclick="loadAdminData()">🔄 Refresh</button>\n' +
'        </div>\n' +
'        <div class="table-responsive">\n' +
'          <table>\n' +
'            <thead><tr><th>UID</th><th>Name</th><th>Department</th><th>Address</th><th>Base Salary</th><th>Bonus</th><th>Status</th><th>Actions</th></tr></thead>\n' +
'            <tbody id="activeTableBody"></tbody>\n' +
'          </table>\n' +
'        </div>\n' +
'      </div>\n' +
'      <div id="adminAddTab" style="display:none;">\n' +
'        <h3 id="formHeading">Add New Faculty Record</h3>\n' +
'        <div class="grid-form">\n' +
'          <input type="hidden" id="editUid">\n' +
'          <div class="form-group"><label>Teacher Name</label><input type="text" id="fName"></div>\n' +
'          <div class="form-group"><label>Gender</label><select id="fGender"><option value="Female">Female</option><option value="Male">Male</option></select></div>\n' +
'          <div class="form-group"><label>Base Monthly Salary (₹)</label><input type="number" id="fSalary" value="55000"></div>\n' +
'          <div class="form-group"><label>Department</label><input type="text" id="fDept" placeholder="e.g. Computer Science"></div>\n' +
'          <div class="form-group"><label>Joining Date</label><input type="date" id="fDate"></div>\n' +
'          <div class="form-group"><label>Residential Address</label><input type="text" id="fAddress" placeholder="e.g. LPU Campus, Punjab"></div>\n' +
'          <div class="form-group"><label>Mobile Phone</label><input type="text" id="fPhone" placeholder="10-digit number"></div>\n' +
'          <div class="form-group"><label>Registered Email</label><input type="email" id="fEmail" placeholder="teacher@domain.com"></div>\n' +
'        </div>\n' +
'        <button class="btn-primary" style="max-width:220px;" onclick="saveFacultyRecord()">Save Record</button>\n' +
'      </div>\n' +
'      <div id="adminBulkTab" style="display:none;">\n' +
'        <h3>Bulk Faculty Upload (Excel or CSV)</h3>\n' +
'        <p style="font-size:13px; color:#64748b; margin:6px 0 16px 0;">Upload .xlsx or .csv. Address and all columns are mapped automatically! Existing names are updated in-place (No Duplicates).</p>\n' +
'        <div class="upload-zone" onclick="document.getElementById(\'bulkFileInput\').click()">\n' +
'          <span style="font-size:32px;">📑</span>\n' +
'          <p style="font-size:14px; font-weight:600; color:#334155; margin-top:8px;">Choose Excel / CSV File</p>\n' +
'          <input type="file" id="bulkFileInput" accept=".xlsx,.xls,.csv" style="display:none;" onchange="handleBulkFileSelect(event)">\n' +
'        </div>\n' +
'        <div id="bulkFileStatus" style="margin-top:14px; font-size:13px; font-weight:600; color:#4338ca; display:none;"></div>\n' +
'        <div id="bulkPreviewSection" style="display:none; margin-top:20px;">\n' +
'          <h4>Parsed Preview (<span id="parsedRowCount">0</span> Records Found)</h4>\n' +
'          <div class="table-responsive"><table><thead><tr><th>Name</th><th>Gender</th><th>Salary</th><th>Department</th><th>Joining Date</th><th>Address</th><th>Phone</th><th>Email</th></tr></thead><tbody id="bulkPreviewBody"></tbody></table></div>\n' +
'          <button id="processBulkBtn" class="btn-primary" style="max-width:280px; margin-top:16px;" onclick="submitBulkUpload()">Process Bulk Import & Sync</button>\n' +
'        </div>\n' +
'      </div>\n' +
'      <div id="adminBinTab" style="display:none;">\n' +
'        <h3>Recycle Bin (BIN RECORDS)</h3>\n' +
'        <p style="font-size:12px; color:#64748b; margin-bottom:12px;">Records purge automatically after 30 seconds unless restored.</p>\n' +
'        <div class="table-responsive"><table><thead><tr><th>UID</th><th>Name</th><th>Dept</th><th>Address</th><th>Deleted At</th><th>Status</th><th>Actions</th></tr></thead><tbody id="binTableBody"></tbody></table></div>\n' +
'      </div>\n' +
'      <div id="adminRequestsTab" style="display:none;">\n' +
'        <h3>Faculty Support Requests (CONTACT US)</h3>\n' +
'        <p style="font-size:12px; color:#64748b; margin-bottom:12px;">Marking Resolved or Rejected notifies faculty and deletes record after 30s.</p>\n' +
'        <div class="table-responsive"><table><thead><tr><th>Req ID</th><th>Date</th><th>UID</th><th>Name</th><th>Type</th><th>Description</th><th>Attachment</th><th>Status</th><th>Actions</th></tr></thead><tbody id="requestsTableBody"></tbody></table></div>\n' +
'      </div>\n' +
'    </div>\n' +
'    <div id="staffSection" class="card" style="display:none;">\n' +
'      <h2 style="font-size:22px; font-weight:800; color:#1e1b4b;" id="staffGreeting">Good Morning</h2>\n' +
'      <p style="font-size:13px; color:#64748b; margin-top:2px;">Verified Faculty Portal & Self-Service Desk</p>\n' +
'      <div class="kpi-row">\n' +
'        <div class="kpi-box"><div class="kpi-title">Monthly Base Pay</div><div class="kpi-val" id="kpiSalary">₹0</div></div>\n' +
'        <div class="kpi-box"><div class="kpi-title">Experience Bonus</div><div class="kpi-val" style="color:#15803d;" id="kpiBonus">+₹0</div></div>\n' +
'        <div class="kpi-box" style="background:#1e1b4b; color:#fff;"><div class="kpi-title" style="color:#c7d2fe;">Total Gross Pay</div><div class="kpi-val" style="color:#fff;" id="kpiTotal">₹0</div></div>\n' +
'      </div>\n' +
'      <div class="profile-card">\n' +
'        <h4 style="margin-bottom:14px; color:#475569;">Profile Details & Quick Contact Update</h4>\n' +
'        <div class="grid-form">\n' +
'          <div class="form-group"><label>Faculty UID (Locked)</label><input type="text" id="sUid" disabled style="background:#f1f5f9;"></div>\n' +
'          <div class="form-group"><label>Department (Locked)</label><input type="text" id="sDept" disabled style="background:#f1f5f9;"></div>\n' +
'          <div class="form-group"><label>Joining Date (Locked)</label><input type="text" id="sDate" disabled style="background:#f1f5f9;"></div>\n' +
'          <div class="form-group"><label>Residential Address (Locked)</label><input type="text" id="sAddress" disabled style="background:#f1f5f9;"></div>\n' +
'          <div class="form-group"><label>Status</label><input type="text" id="sStatus" disabled style="background:#f1f5f9; color:#15803d; font-weight:bold;"></div>\n' +
'          <div class="form-group"><label>Registered Mobile (Editable)</label><input type="text" id="sPhone"></div>\n' +
'          <div class="form-group"><label>Registered Email (Editable)</label><input type="email" id="sEmail"></div>\n' +
'        </div>\n' +
'        <button class="btn-primary" style="max-width:260px; margin-top:10px;" onclick="updateStaffContact()">Update Mobile & Email</button>\n' +
'      </div>\n' +
'      <div class="profile-card" style="margin-top:25px;">\n' +
'        <h4 style="margin-bottom:14px; color:#475569;">Submit Service / Support Request (CONTACT US)</h4>\n' +
'        <div class="form-group">\n' +
'          <label>Request Type</label>\n' +
'          <select id="reqType">\n' +
'            <option value="Profile Update">Profile Update</option>\n' +
'            <option value="Salary / Bonus Query">Salary / Bonus Query</option>\n' +
'            <option value="Leave / Attendance Issue">Leave / Attendance Issue</option>\n' +
'            <option value="Document Verification">Document Verification</option>\n' +
'            <option value="Technical Support">Technical Support</option>\n' +
'            <option value="Other">Other</option>\n' +
'          </select>\n' +
'        </div>\n' +
'        <div class="form-group"><label>Description</label><textarea id="reqDesc" rows="3" placeholder="Describe query..."></textarea></div>\n' +
'        <div class="form-group"><label>Attachment (PDF / JPG / ZIP < 10 MB)</label><input type="file" id="reqFile" accept=".pdf,.jpg,.jpeg,.zip"></div>\n' +
'        <button class="btn-primary" style="max-width:240px;" onclick="submitSupportRequest()">Submit Ticket</button>\n' +
'      </div>\n' +
'    </div>\n' +
'  </div>\n' +
'  <footer><p>© Lovely Professional University • Directorate of Human Capital</p></footer>\n' +
'  <script>\n' +
'    var SCRIPT_URL = "https://script.google.com/macros/s/AKfycby6T-o1_u1nc3-y-QYfcsmw202waGk1F5gxB3yG7FdMlc-qeVh6lkwBcRJY8QZa6ONu6Q/exec";\n' +
'    var activeStaffRecord = null;\n' +
'    var countdownTimer = null;\n' +
'    var timeLeft = 120;\n' +
'    var parsedBulkRows = [];\n' +
'    function updateLiveClock() {\n' +
'      var now = new Date();\n' +
'      var timeStr = now.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });\n' +
'      document.getElementById("liveClockDisplay").innerText = "🕒 " + timeStr;\n' +
'      var hrs = now.getHours();\n' +
'      var greeting = "Good Morning";\n' +
'      if (hrs >= 12 && hrs < 17) greeting = "Good Afternoon";\n' +
'      else if (hrs >= 17) greeting = "Good Evening";\n' +
'      if (activeStaffRecord) document.getElementById("staffGreeting").innerText = greeting + ", " + activeStaffRecord.name;\n' +
'    }\n' +
'    setInterval(updateLiveClock, 1000);\n' +
'    updateLiveClock();\n' +
'    function callBackend(actionName, payload) {\n' +
'      payload = payload || {};\n' +
'      return new Promise(function(resolve, reject) {\n' +
'        if (typeof google !== "undefined" && google.script && google.script.run) {\n' +
'          google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).handleApiRequest(actionName, JSON.stringify(payload));\n' +
'        } else {\n' +
'          fetch(SCRIPT_URL, {\n' +
'            method: "POST",\n' +
'            headers: { "Content-Type": "text/plain;charset=utf-8" },\n' +
'            body: JSON.stringify(Object.assign({ action: actionName }, payload))\n' +
'          }).then(function(r){ return r.json(); }).then(resolve).catch(reject);\n' +
'        }\n' +
'      });\n' +
'    }\n' +
'    function switchRoleFields() {\n' +
'      var role = document.getElementById("loginRole").value;\n' +
'      document.getElementById("adminFields").style.display = role === "ADMIN" ? "block" : "none";\n' +
'      document.getElementById("staffFields").style.display = role === "STAFF" ? "block" : "none";\n' +
'      document.getElementById("otpStep").style.display = "none";\n' +
'      document.getElementById("loginAlert").style.display = "none";\n' +
'      document.getElementById("loginSubmitBtn").innerText = "Continue";\n' +
'      clearInterval(countdownTimer);\n' +
'    }\n' +
'    function showAlert(msg, isSuccess) {\n' +
'      var b = document.getElementById("loginAlert");\n' +
'      b.className = "alert-box " + (isSuccess ? "alert-success" : "alert-error");\n' +
'      b.innerText = msg; b.style.display = "block";\n' +
'    }\n' +
'    function startOtpCountdown() {\n' +
'      clearInterval(countdownTimer);\n' +
'      timeLeft = 120;\n' +
'      var timerDisplay = document.getElementById("timerDisplay");\n' +
'      var resendBtn = document.getElementById("resendOtpBtn");\n' +
'      resendBtn.disabled = true; resendBtn.classList.remove("active");\n' +
'      function update() {\n' +
'        var mins = Math.floor(timeLeft / 60);\n' +
'        var secs = timeLeft % 60;\n' +
'        timerDisplay.innerText = "⏱️ Expires in: " + (mins < 10 ? "0" : "") + mins + ":" + (secs < 10 ? "0" : "") + secs;\n' +
'        if (timeLeft <= 0) {\n' +
'          clearInterval(countdownTimer);\n' +
'          timerDisplay.innerText = "⚠️ OTP Expired";\n' +
'          resendBtn.disabled = false; resendBtn.classList.add("active");\n' +
'        }\n' +
'        timeLeft--;\n' +
'      }\n' +
'      update(); countdownTimer = setInterval(update, 1000);\n' +
'    }\n' +
'    function handleResendOtp() {\n' +
'      var id = activeStaffRecord ? activeStaffRecord.uid : document.getElementById("staffIdentifier").value;\n' +
'      var jDate = document.getElementById("staffJoiningDate").value;\n' +
'      showAlert("Sending fresh OTP...", true);\n' +
'      document.getElementById("staffOtpInput").value = "";\n' +
'      callBackend("apiStaffRequestOtp", { identifier: id, joiningDate: jDate }).then(function(res){\n' +
'        if (res.success) { showAlert(res.message, true); startOtpCountdown(); }\n' +
'        else showAlert(res.message, false);\n' +
'      });\n' +
'    }\n' +
'    function handleLoginSubmit() {\n' +
'      var role = document.getElementById("loginRole").value;\n' +
'      var btn = document.getElementById("loginSubmitBtn");\n' +
'      if (role === "ADMIN") {\n' +
'        var u = document.getElementById("adminUser").value.trim();\n' +
'        var p = document.getElementById("adminPass").value.trim();\n' +
'        btn.disabled = true;\n' +
'        callBackend("apiAdminLogin", { username: u, password: p }).then(function(res){\n' +
'          btn.disabled = false;\n' +
'          if (res.success) {\n' +
'            document.getElementById("loginSection").style.display = "none";\n' +
'            document.getElementById("adminSection").style.display = "block";\n' +
'            document.getElementById("logoutBtn").style.display = "block";\n' +
'            loadAdminData();\n' +
'          } else showAlert(res.message, false);\n' +
'        }).catch(function(err){ btn.disabled = false; showAlert(err.message, false); });\n' +
'      } else {\n' +
'        if (document.getElementById("otpStep").style.display === "none") {\n' +
'          var id = document.getElementById("staffIdentifier").value.trim();\n' +
'          var jDate = document.getElementById("staffJoiningDate").value;\n' +
'          if (!id || !jDate) { showAlert("Please enter your Mobile/UID and Date of Joining.", false); return; }\n' +
'          btn.disabled = true; showAlert("Checking registry & dispatching OTP...", true);\n' +
'          callBackend("apiStaffRequestOtp", { identifier: id, joiningDate: jDate }).then(function(res){\n' +
'            btn.disabled = false;\n' +
'            if (res.success) {\n' +
'              document.getElementById("staffFields").style.display = "none";\n' +
'              document.getElementById("otpStep").style.display = "block";\n' +
'              btn.innerText = "Verify OTP & Sign In"; showAlert(res.message, true); startOtpCountdown();\n' +
'            } else showAlert(res.message, false);\n' +
'          }).catch(function(err){ btn.disabled = false; showAlert(err.message, false); });\n' +
'        } else {\n' +
'          var otp = document.getElementById("staffOtpInput").value.trim();\n' +
'          var sId = document.getElementById("staffIdentifier").value.trim();\n' +
'          btn.disabled = true;\n' +
'          callBackend("apiStaffVerifyOtp", { uid: sId, enteredOtp: otp }).then(function(res){\n' +
'            btn.disabled = false;\n' +
'            if (res.success) {\n' +
'              clearInterval(countdownTimer); activeStaffRecord = res.record;\n' +
'              document.getElementById("loginSection").style.display = "none";\n' +
'              document.getElementById("staffSection").style.display = "block";\n' +
'              document.getElementById("logoutBtn").style.display = "block";\n' +
'              renderStaffProfile(res.record);\n' +
'            } else showAlert(res.message, false);\n' +
'          }).catch(function(err){ btn.disabled = false; showAlert(err.message, false); });\n' +
'        }\n' +
'      }\n' +
'    }\n' +
'    function renderStaffProfile(r) {\n' +
'      updateLiveClock();\n' +
'      document.getElementById("kpiSalary").innerText = "₹" + r.salary.toLocaleString("en-IN");\n' +
'      document.getElementById("kpiBonus").innerText = "+₹" + r.bonus.toLocaleString("en-IN");\n' +
'      document.getElementById("kpiTotal").innerText = "₹" + r.total.toLocaleString("en-IN");\n' +
'      document.getElementById("sUid").value = r.uid;\n' +
'      document.getElementById("sDept").value = r.dept;\n' +
'      document.getElementById("sDate").value = r.joiningDate;\n' +
'      document.getElementById("sAddress").value = r.address || "N/A";\n' +
'      document.getElementById("sStatus").value = r.status.toUpperCase();\n' +
'      document.getElementById("sPhone").value = r.phone;\n' +
'      document.getElementById("sEmail").value = r.email;\n' +
'    }\n' +
'    function updateStaffContact() {\n' +
'      var p = document.getElementById("sPhone").value.trim();\n' +
'      var e = document.getElementById("sEmail").value.trim();\n' +
'      callBackend("apiStaffUpdateContact", { uid: activeStaffRecord.uid, phone: p, email: e }).then(function(res){\n' +
'        alert(res.message);\n' +
'      });\n' +
'    }\n' +
'    function submitSupportRequest() {\n' +
'      var type = document.getElementById("reqType").value;\n' +
'      var desc = document.getElementById("reqDesc").value.trim();\n' +
'      var fi = document.getElementById("reqFile");\n' +
'      if (!desc) { alert("Please enter description."); return; }\n' +
'      var payload = {\n' +
'        uid: activeStaffRecord.uid, name: activeStaffRecord.name, phone: document.getElementById("sPhone").value, email: document.getElementById("sEmail").value,\n' +
'        type: type, description: desc, fileBase64: "", fileName: "", fileType: ""\n' +
'      };\n' +
'      if (fi.files.length > 0) {\n' +
'        var f = fi.files[0];\n' +
'        if (f.size > 10 * 1024 * 1024) { alert("File exceeds 10 MB limit."); return; }\n' +
'        var r = new FileReader();\n' +
'        r.onload = function(evt) {\n' +
'          payload.fileBase64 = evt.target.result;\n' +
'          payload.fileName = f.name; payload.fileType = f.type;\n' +
'          callBackend("apiStaffSubmitRequest", { data: payload }).then(function(res){ alert(res.message); fi.value = ""; document.getElementById("reqDesc").value = ""; });\n' +
'        };\n' +
'        r.readAsDataURL(f);\n' +
'      } else {\n' +
'        callBackend("apiStaffSubmitRequest", { data: payload }).then(function(res){ alert(res.message); document.getElementById("reqDesc").value = ""; });\n' +
'      }\n' +
'    }\n' +
'    function logout() { location.reload(); }\n' +
'    function switchAdminTab(tab, btn) {\n' +
'      document.querySelectorAll(".tab-btn").forEach(function(b){ b.classList.remove("active"); });\n' +
'      btn.classList.add("active");\n' +
'      document.getElementById("adminActiveTab").style.display = tab === "ACTIVE_TAB" ? "block" : "none";\n' +
'      document.getElementById("adminAddTab").style.display = tab === "ADD_TAB" ? "block" : "none";\n' +
'      document.getElementById("adminBulkTab").style.display = tab === "BULK_TAB" ? "block" : "none";\n' +
'      document.getElementById("adminBinTab").style.display = tab === "BIN_TAB" ? "block" : "none";\n' +
'      document.getElementById("adminRequestsTab").style.display = tab === "REQUESTS_TAB" ? "block" : "none";\n' +
'    }\n' +
'    var allFacultyRecords = [];\n' +
'    function loadAdminData() {\n' +
'      callBackend("apiAdminGetAllData").then(function(data){\n' +
'        allFacultyRecords = data.active;\n' +
'        var b = document.getElementById("activeTableBody"); b.innerHTML = "";\n' +
'        data.active.forEach(function(r){\n' +
'          b.innerHTML += "<tr><td><code>" + r.uid + "</code></td><td><strong>" + r.name + "</strong></td><td>" + r.dept + "</td><td>" + (r.address || "N/A") + "</td><td>₹" + r.salary.toLocaleString("en-IN") + "</td><td>₹" + r.bonus.toLocaleString("en-IN") + "</td><td><span style=\'color:#15803d; font-weight:bold;\'>● " + r.status + "</span></td><td><button class=\'btn-sm btn-edit\' onclick=\'editRecord(\\"" + r.uid + "\\")\'>Edit</button><button class=\'btn-sm btn-bin\' onclick=\'binRecord(\\"" + r.uid + "\\")\'>Move to Bin</button></td></tr>";\n' +
'        });\n' +
'        var binB = document.getElementById("binTableBody"); binB.innerHTML = "";\n' +
'        data.bin.forEach(function(r){\n' +
'          binB.innerHTML += "<tr><td><code>" + r.uid + "</code></td><td>" + r.name + "</td><td>" + r.dept + "</td><td>" + (r.address || "N/A") + "</td><td>" + r.deletedAt + "</td><td><span style=\'color:#b91c1c; font-weight:bold;\'>● Deleted</span></td><td><button class=\'btn-sm btn-restore\' onclick=\'restoreRecord(\\"" + r.uid + "\\")\'>Restore</button></td></tr>";\n' +
'        });\n' +
'        var reqB = document.getElementById("requestsTableBody"); reqB.innerHTML = "";\n' +
'        var pendingCount = 0;\n' +
'        data.requests.forEach(function(rq){\n' +
'          if (rq.status === "Pending") pendingCount++;\n' +
'          var link = rq.attachUrl !== "No Attachment" ? "<a href=\'" + rq.attachUrl + "\' target=\'_blank\'>View File</a>" : "None";\n' +
'          reqB.innerHTML += "<tr><td><strong>" + rq.reqId + "</strong></td><td>" + rq.timestamp + "</td><td><code>" + rq.uid + "</code></td><td>" + rq.name + "</td><td>" + rq.type + "</td><td>" + rq.desc + "</td><td>" + link + "</td><td><strong>" + rq.status + "</strong></td><td><button class=\'btn-sm btn-resolve\' onclick=\'updateReqStatus(\\"" + rq.reqId + "\\", \\"Resolved\\")\'>Resolve</button><button class=\'btn-sm btn-reject\' onclick=\'updateReqStatus(\\"" + rq.reqId + "\\", \\"Rejected\\")\'>Reject</button></td></tr>";\n' +
'        });\n' +
'        document.getElementById("pendingBadge").innerText = pendingCount;\n' +
'      });\n' +
'    }\n' +
'    function editRecord(uid) {\n' +
'      var rec = allFacultyRecords.find(function(r){ return r.uid === uid; });\n' +
'      if (!rec) return;\n' +
'      document.getElementById("editUid").value = rec.uid;\n' +
'      document.getElementById("fName").value = rec.name;\n' +
'      document.getElementById("fGender").value = rec.gender;\n' +
'      document.getElementById("fSalary").value = rec.salary;\n' +
'      document.getElementById("fDept").value = rec.dept;\n' +
'      document.getElementById("fDate").value = rec.joiningDate;\n' +
'      document.getElementById("fAddress").value = rec.address || "";\n' +
'      document.getElementById("fPhone").value = rec.phone;\n' +
'      document.getElementById("fEmail").value = rec.email;\n' +
'      document.getElementById("formHeading").innerText = "Modify Record: " + rec.uid;\n' +
'      switchAdminTab("ADD_TAB", document.querySelectorAll(".tab-btn")[1]);\n' +
'    }\n' +
'    function saveFacultyRecord() {\n' +
'      var editUid = document.getElementById("editUid").value;\n' +
'      var payload = {\n' +
'        uid: editUid, name: document.getElementById("fName").value, gender: document.getElementById("fGender").value, salary: document.getElementById("fSalary").value,\n' +
'        dept: document.getElementById("fDept").value, joiningDate: document.getElementById("fDate").value, address: document.getElementById("fAddress").value, phone: document.getElementById("fPhone").value, email: document.getElementById("fEmail").value\n' +
'      };\n' +
'      if (!payload.name || !payload.email) { alert("Name and Email required."); return; }\n' +
'      if (editUid) {\n' +
'        callBackend("apiAdminUpdateRecord", { record: payload }).then(function(){ alert("Updated!"); document.getElementById("editUid").value = ""; loadAdminData(); switchAdminTab("ACTIVE_TAB", document.querySelectorAll(".tab-btn")[0]); });\n' +
'      } else {\n' +
'        callBackend("apiAdminAddRecord", { record: payload }).then(function(res){ alert("Added UID: " + res.uid); loadAdminData(); switchAdminTab("ACTIVE_TAB", document.querySelectorAll(".tab-btn")[0]); });\n' +
'      }\n' +
'    }\n' +
'    function handleBulkFileSelect(e) {\n' +
'      var file = e.target.files[0];\n' +
'      if (!file) return;\n' +
'      var statusBox = document.getElementById("bulkFileStatus");\n' +
'      statusBox.style.display = "block"; statusBox.innerText = "Analyzing columns in " + file.name + "...";\n' +
'      var reader = new FileReader();\n' +
'      reader.onload = function(evt) {\n' +
'        try {\n' +
'          var data = new Uint8Array(evt.target.result);\n' +
'          var workbook = XLSX.read(data, { type: "array", cellDates: true });\n' +
'          var sheet = workbook.Sheets[workbook.SheetNames[0]];\n' +
'          var rawJson = XLSX.utils.sheet_to_json(sheet, { header: 1 });\n' +
'          if (!rawJson || rawJson.length <= 1) {\n' +
'            alert("File contains no data rows.");\n' +
'            statusBox.style.display = "none";\n' +
'            return;\n' +
'          }\n' +
'          var headerRow = rawJson[0] || [];\n' +
'          var nameIdx = -1, genderIdx = -1, salaryIdx = -1, deptIdx = -1, dateIdx = -1, addrIdx = -1, phoneIdx = -1, emailIdx = -1;\n' +
'          for (var c = 0; c < headerRow.length; c++) {\n' +
'            var h = String(headerRow[c] || "").toLowerCase().trim();\n' +
'            if (nameIdx === -1 && (h.includes("name") || h.includes("teacher") || h.includes("faculty"))) nameIdx = c;\n' +
'            else if (genderIdx === -1 && (h.includes("gender") || h.includes("sex"))) genderIdx = c;\n' +

'            else if (salaryIdx === -1 && (h.includes("salary") || h.includes("pay") || h.includes("ctc") || h.includes("remuneration"))) salaryIdx = c;\n' +
'            else if (deptIdx === -1 && (h.includes("dept") || h.includes("department") || h.includes("subject") || h.includes("branch"))) deptIdx = c;\n' +
'            else if (dateIdx === -1 && (h.includes("join") || h.includes("doj") || h.includes("date") || h.includes("appointment"))) dateIdx = c;\n' +
'            else if (addrIdx === -1 && (h.includes("address") || h.includes("addr") || h.includes("location") || h.includes("residence"))) addrIdx = c;\n' +
'            else if (phoneIdx === -1 && (h.includes("phone") || h.includes("mobile") || h.includes("contact") || h.includes("cell"))) phoneIdx = c;\n' +
'            else if (emailIdx === -1 && (h.includes("email") || h.includes("mail"))) emailIdx = c;\n' +
'          }\n' +
'          var startRow = 1;\n' +
'          if (nameIdx === -1) {\n' +
'            startRow = 0;\n' +
'            nameIdx = 0; genderIdx = 1; salaryIdx = 2; deptIdx = 3; dateIdx = 4; addrIdx = 5; phoneIdx = 6; emailIdx = 7;\n' +
'          }\n' +
'          parsedBulkRows = [];\n' +
'          for (var r = startRow; r < rawJson.length; r++) {\n' +
'            var row = rawJson[r];\n' +
'            if (!row || row.length === 0) continue;\n' +
'            var rawName = String(row[nameIdx] || "").trim();\n' +
'            if (!rawName) continue;\n' +
'            var rawGender = genderIdx !== -1 ? String(row[genderIdx] || "Female").trim() : "Female";\n' +
'            var rawSalary = salaryIdx !== -1 ? Number(String(row[salaryIdx] || "").replace(/[^0-9.]/g, "")) || 50000 : 50000;\n' +
'            var rawDept = deptIdx !== -1 ? String(row[deptIdx] || "General Faculty").trim() : "General Faculty";\n' +
'            var rawAddr = addrIdx !== -1 ? String(row[addrIdx] || "").trim() : "";\n' +
'            var rawDate = dateIdx !== -1 ? row[dateIdx] : new Date();\n' +
'            var formattedDate = "2020-01-01";\n' +
'            if (rawDate instanceof Date && !isNaN(rawDate)) {\n' +
'              formattedDate = rawDate.toISOString().split("T")[0];\n' +
'            } else if (typeof rawDate === "string" && rawDate.trim()) {\n' +
'              var parts = rawDate.trim().split(/[-/]/);\n' +
'              if (parts.length === 3) {\n' +
'                if (parts[0].length === 4) formattedDate = parts[0] + "-" + parts[1].padStart(2, "0") + "-" + parts[2].padStart(2, "0");\n' +
'                else formattedDate = parts[2] + "-" + parts[1].padStart(2, "0") + "-" + parts[0].padStart(2, "0");\n' +
'              }\n' +
'            }\n' +
'            var rawPhone = phoneIdx !== -1 ? String(row[phoneIdx] || "").trim() : "";\n' +
'            var rawEmail = emailIdx !== -1 ? String(row[emailIdx] || "").trim() : "";\n' +
'            parsedBulkRows.push({\n' +
'              name: rawName, gender: rawGender, salary: rawSalary, dept: rawDept, joiningDate: formattedDate, address: rawAddr, phone: rawPhone, email: rawEmail\n' +
'            });\n' +
'          }\n' +
'          renderBulkPreview();\n' +
'          statusBox.innerText = "✅ " + file.name + " parsed successfully (" + parsedBulkRows.length + " records mapped)";\n' +
'        } catch (err) {\n' +
'          alert("Error parsing file: " + err.message);\n' +
'          statusBox.style.display = "none";\n' +
'        }\n' +
'      };\n' +
'      reader.readAsArrayBuffer(file);\n' +
'    }\n' +
'    function renderBulkPreview() {\n' +
'      var b = document.getElementById("bulkPreviewBody"); b.innerHTML = "";\n' +
'      document.getElementById("parsedRowCount").innerText = parsedBulkRows.length;\n' +
'      var limit = Math.min(parsedBulkRows.length, 10);\n' +
'      for (var i = 0; i < limit; i++) {\n' +
'        var it = parsedBulkRows[i];\n' +
'        b.innerHTML += "<tr><td><strong>" + it.name + "</strong></td><td>" + it.gender + "</td><td>₹" + it.salary.toLocaleString("en-IN") + "</td><td>" + it.dept + "</td><td>" + it.joiningDate + "</td><td>" + (it.address || "N/A") + "</td><td>" + it.phone + "</td><td>" + it.email + "</td></tr>";\n' +
'      }\n' +
'      if (parsedBulkRows.length > 10) {\n' +
'        b.innerHTML += "<tr><td colspan=\'8\' style=\'text-align:center; color:#64748b;\'><em>...and " + (parsedBulkRows.length - 10) + " more records ready</em></td></tr>";\n' +
'      }\n' +
'      document.getElementById("bulkPreviewSection").style.display = "block";\n' +
'    }\n' +
'    function submitBulkUpload() {\n' +
'      var btn = document.getElementById("processBulkBtn");\n' +
'      btn.disabled = true; btn.innerText = "Processing Bulk Import (Please wait)...";\n' +
'      callBackend("apiAdminBulkUpload", { records: parsedBulkRows }).then(function(res){\n' +
'        btn.disabled = false; btn.innerText = "Process Bulk Import & Sync";\n' +
'        alert(res.message); document.getElementById("bulkPreviewSection").style.display = "none";\n' +
'        document.getElementById("bulkFileStatus").style.display = "none";\n' +
'        document.getElementById("bulkFileInput").value = ""; parsedBulkRows = []; loadAdminData(); switchAdminTab("ACTIVE_TAB", document.querySelectorAll(".tab-btn")[0]);\n' +
'      }).catch(function(err){ btn.disabled = false; btn.innerText = "Process Bulk Import & Sync"; alert("Error: " + err.message); });\n' +
'    }\n' +
'    function binRecord(uid) {\n' +
'      if (confirm("Move " + uid + " to Bin?")) callBackend("apiAdminMoveToBin", { uid: uid }).then(loadAdminData);\n' +
'    }\n' +
'    function restoreRecord(uid) {\n' +
'      callBackend("apiAdminRestoreRecord", { uid: uid }).then(function(){ alert("Restored!"); loadAdminData(); });\n' +
'    }\n' +
'    function updateReqStatus(reqId, status) {\n' +
'      if (confirm("Mark " + reqId + " as " + status + "?")) {\n' +
'        callBackend("apiAdminUpdateRequestStatus", { reqId: reqId, newStatus: status }).then(loadAdminData);\n' +
'      }\n' +
'    }\n' +
'  </script>\n' +
'</body>\n' +
'</html>';
}