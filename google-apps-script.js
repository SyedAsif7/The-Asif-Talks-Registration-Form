/**
 * Google Apps Script Backend for "The Asif Talks (Episode #2)" Registration
 * Distinguished Guest: Hon. Mayor Syed Iqbal Syed Khwaja (Mayor, Parbhani Municipal Corporation)
 *
 * ---------------------------------------------------------------------------------------------------------------------------------
 * 📋 GOOGLE SHEET HEADERS (ROW 1):
 * Column A: Timestamp
 * Column B: Pass ID
 * Column C: Student Name
 * Column D: Mobile Number
 * Column E: Email Address
 * Column F: Gender
 * Column G: College / Institute / Organization
 * Column H: City / Location
 * Column I: How Heard
 * Column J: Question for Hon. Mayor Syed Iqbal
 * Column K: Photo & Video Consent
 * ---------------------------------------------------------------------------------------------------------------------------------
 */

// Helper to get or create the Episode 2 sheet tab
function getEpisodeSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var primarySheetName = "The-Asif-Talks-Registration-Form-Episode-2";
  var fallbackSheetName = "Episode 2";
  
  var sheet = ss.getSheetByName(primarySheetName) || ss.getSheetByName(fallbackSheetName);
  
  if (!sheet) {
    var activeSheet = ss.getActiveSheet();
    if (activeSheet.getLastRow() === 0) {
      sheet = activeSheet;
      sheet.setName(primarySheetName);
    } else {
      sheet = ss.insertSheet(primarySheetName);
    }
    
    // Set headers
    var headers = [
      "Timestamp",
      "Pass ID",
      "Student Name",
      "Mobile Number",
      "Email Address",
      "Gender",
      "College / Institute / Organization",
      "City / Location",
      "How Heard",
      "Question for Hon. Mayor Syed Iqbal",
      "Photo & Video Consent"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0f2744").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Helper to normalize phone number to last 10 digits
function normalizePhone(num) {
  if (!num) return "";
  var cleaned = String(num).replace(/[^0-9]/g, "");
  if (cleaned.length > 10) {
    cleaned = cleaned.slice(-10); // Take last 10 digits
  }
  return cleaned;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevent concurrent write collisions

  try {
    var sheet = getEpisodeSheet();
    var lastRow = sheet.getLastRow();

    // Extract POST parameters
    var p = e.parameter || {};
    var name = (p.name || "").trim();
    var rawPhone = (p.phone || "").trim();
    var cleanPhone = normalizePhone(rawPhone);
    var email = (p.email || "").trim();
    var gender = (p.gender || "").trim();
    var college = (p.college || "").trim();
    var city = (p.city || "").trim();
    var source = (p.source || "").trim();
    var question = (p.question || "").trim();
    var consent = p.consent || "Agreed";

    // Check for duplicate mobile number in Column D (Rows 2 to lastRow)
    if (cleanPhone && cleanPhone.length === 10 && lastRow > 1) {
      var phoneValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
      var passIdValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      
      for (var i = 0; i < phoneValues.length; i++) {
        var existingClean = normalizePhone(phoneValues[i][0]);
        if (existingClean && existingClean === cleanPhone) {
          var existingPassId = passIdValues[i][0] || ("TAT-" + ("000" + (i + 1)).slice(-3));
          return ContentService
            .createTextOutput(JSON.stringify({ 
              result: "duplicate",
              error: "DUPLICATE_PHONE",
              passId: existingPassId,
              message: "This mobile number is already registered for The Asif Talks (Episode #2)! Pass ID: " + existingPassId 
            }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // Auto-calculate sequential Pass ID (TAT-001, TAT-002, etc.)
    var passCount = Math.max(1, lastRow); // Row 1 is header
    var passId = "TAT-" + ("000" + passCount).slice(-3);

    var rowData = [
      new Date(),
      passId,
      name,
      rawPhone,
      email,
      gender,
      college,
      city,
      source,
      question,
      consent
    ];

    sheet.appendRow(rowData);

    var totalSeats = 200;
    var remainingSeats = Math.max(0, totalSeats - passCount);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        result: "success", 
        passId: passId, 
        registeredCount: passCount,
        remainingSeats: remainingSeats,
        message: "Registration for Episode #2 recorded successfully" 
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        result: "error", 
        error: err.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var sheet = getEpisodeSheet();
  var lastRow = sheet.getLastRow();
  var count = Math.max(0, lastRow - 1);
  var totalSeats = 200;
  var remainingSeats = Math.max(0, totalSeats - count);
  var nextPassId = "TAT-" + ("000" + (count + 1)).slice(-3);

  // Check if a specific mobile number already exists & collect all registered phones for Ep 2
  var phoneList = [];
  var checkPhone = e && e.parameter && (e.parameter.checkPhone || e.parameter.phone) ? normalizePhone(e.parameter.checkPhone || e.parameter.phone) : "";
  var isDuplicate = false;
  var duplicatePassId = "";

  if (lastRow > 1) {
    var phoneValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    var passIdValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

    for (var i = 0; i < phoneValues.length; i++) {
      var existingClean = normalizePhone(phoneValues[i][0]);
      if (existingClean) {
        phoneList.push(existingClean);
        if (checkPhone && existingClean === checkPhone) {
          isDuplicate = true;
          duplicatePassId = passIdValues[i][0] || ("TAT-" + ("000" + (i + 1)).slice(-3));
        }
      }
    }
  }

  var responseData = {
    status: "live",
    episode: "Episode #2",
    guest: "Hon. Mayor Syed Iqbal Syed Khwaja",
    totalSeats: totalSeats,
    registeredCount: count,
    remainingSeats: remainingSeats,
    nextPassId: nextPassId,
    isDuplicate: isDuplicate,
    duplicatePassId: duplicatePassId,
    registeredPhones: phoneList
  };

  return ContentService
    .createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this test function in Apps Script to verify permissions and initialize the sheet tab!
 */
function testSetup() {
  var sheet = getEpisodeSheet();
  Logger.log("Episode 2 Sheet successfully connected: " + sheet.getName());
}
