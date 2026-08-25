/**
 * Google Apps Script Backend for "The Asif Talks (Episode #1)" Registration
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
 * Column J: Question for Guest Speaker
 * Column K: Photo & Video Consent
 * ---------------------------------------------------------------------------------------------------------------------------------
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevent concurrent write collisions

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Auto-calculate sequential Pass ID (TAT-001, TAT-002, etc.)
    var lastRow = sheet.getLastRow();
    var passCount = Math.max(1, lastRow); // Row 1 is header
    var passId = "TAT-" + ("000" + passCount).slice(-3);

    // Extract POST parameters
    var p = e.parameter || {};
    var name = p.name || "";
    var phone = p.phone || "";
    var email = p.email || "";
    var gender = p.gender || "";
    var college = p.college || "";
    var city = p.city || "";
    var source = p.source || "";
    var question = p.question || "";
    var consent = p.consent || "Agreed";

    var rowData = [
      new Date(),
      passId,
      name,
      phone,
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
        message: "Registration recorded successfully" 
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var count = Math.max(0, lastRow - 1);
  var totalSeats = 200;
  var remainingSeats = Math.max(0, totalSeats - count);
  var nextPassId = "TAT-" + ("000" + (count + 1)).slice(-3);

  var responseData = {
    status: "live",
    totalSeats: totalSeats,
    registeredCount: count,
    remainingSeats: remainingSeats,
    nextPassId: nextPassId
  };

  return ContentService
    .createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}
