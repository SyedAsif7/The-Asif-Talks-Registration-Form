/**
 * Google Apps Script Backend for "The Asif Talks (Episode #1)" Registration
 *
 * ---------------------------------------------------------------------------------------------------------------------------------
 * 📋 GOOGLE SHEET SETUP (ROW 1 HEADERS):
 * Column A: Timestamp
 * Column B: Pass ID
 * Column C: Student Name
 * Column D: Mobile Number
 * Column E: Email Address
 * Column F: Gender
 * Column G: College / Institute / Organization
 * Column H: City / Location
 * Column I: Expectations
 * Column J: How Heard
 * Column K: Question for Collector
 * Column L: Photo/Video Consent
 * ---------------------------------------------------------------------------------------------------------------------------------
 *
 * 🚀 DEPLOYMENT INSTRUCTIONS:
 * 1. Open your Google Sheet -> Extensions -> Apps Script.
 * 2. Paste this code into Code.gs and click Save (Ctrl+S).
 * 3. Click "Deploy" -> "New deployment".
 * 4. Select type: "Web app".
 * 5. Configuration:
 *    - Description: The Asif Talks v2 (Full Form)
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 6. Click "Deploy", authorize access, and copy the Web App URL.
 * 7. Paste the Web App URL into line 399 of index.html.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // 10-second timeout to handle peak concurrent registrations

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Auto-calculate sequential Pass ID (e.g. TAT-001, TAT-002)
    var lastRow = sheet.getLastRow();
    var passId = "TAT-" + ("000" + (lastRow)).slice(-3);

    // Extract POST parameters
    var p = e.parameter || {};
    var name = p.name || "";
    var phone = p.phone || "";
    var email = p.email || "";
    var gender = p.gender || "";
    var college = p.college || "";
    var city = p.city || "";
    var expectations = p.expectations || "";
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
      expectations,
      source,
      question,
      consent
    ];

    sheet.appendRow(rowData);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        result: "success", 
        passId: passId, 
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
  return ContentService
    .createTextOutput("The Asif Talks Registration API is operational.")
    .setMimeType(ContentService.MimeType.TEXT);
}
