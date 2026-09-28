/**
 * Google Apps Script Backend for "The Asif Talks (Episode #3)" Registration & Gate Entry Check-in
 * Distinguished Guest: Hon. Smt. Meghana Sakore-Bordikar
 * (Minister of State, Government of Maharashtra | Guardian Minister, Parbhani District)
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
 * Column J: Question for Hon. Smt. Meghana Bordikar
 * Column K: Photo & Video Consent
 * Column L: Check-in Status
 * Column M: Check-in Time
 * Column N: Gate Volunteer
 * ---------------------------------------------------------------------------------------------------------------------------------
 */

// Helper to get or create the Episode 3 sheet tab
function getEpisodeSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var primarySheetName = "The-Asif-Talks-Registration-Form-Episode-3";
  var fallbackSheetName = "Episode 3";
  
  var sheet = ss.getSheetByName(primarySheetName) || ss.getSheetByName(fallbackSheetName);
  
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
    "Question for Hon. Smt. Meghana Bordikar",
    "Photo & Video Consent",
    "Check-in Status",
    "Check-in Time",
    "Gate Volunteer"
  ];

  if (!sheet) {
    var activeSheet = ss.getActiveSheet();
    if (activeSheet.getLastRow() === 0) {
      sheet = activeSheet;
      sheet.setName(primarySheetName);
    } else {
      sheet = ss.insertSheet(primarySheetName);
    }
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0f2744").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  } else {
    // If headers exist, ensure check-in columns (L, M, N) are present
    var lastCol = sheet.getLastColumn();
    if (lastCol < headers.length) {
      for (var col = lastCol + 1; col <= headers.length; col++) {
        sheet.getRange(1, col).setValue(headers[col - 1]).setFontWeight("bold").setBackground("#0f2744").setFontColor("#ffffff");
      }
    }
  }
  return sheet;
}

// Helper to normalize phone number to last 10 digits for strict uniqueness
function normalizePhone(num) {
  if (!num) return "";
  var cleaned = String(num).replace(/[^0-9]/g, "");
  if (cleaned.length >= 10) {
    return cleaned.slice(-10); // Check last 10 digits
  }
  return cleaned;
}

// Automatically send attendee Pass Confirmation Email with Guidelines
function sendConfirmationEmail(email, name, passId) {
  if (!email || email.indexOf("@") === -1) return;
  
  var subject = "🎟️ Your Official Pass for The Asif Talks (Episode #3) — " + passId;
  
  var htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.06);">
      <div style="background: linear-gradient(135deg, #0f2744 0%, #060d17 100%); color: #ffffff; padding: 24px 20px; text-align: center;">
        <h1 style="margin: 0; font-size: 22px; letter-spacing: 0.8px; color: #ffffff;">THE ASIF TALKS • EPISODE #3</h1>
        <p style="margin: 6px 0 0 0; font-size: 13px; color: #93c5fd;">Official Live Studio Audience Boarding Pass</p>
      </div>
      
      <div style="padding: 24px 20px;">
        <p style="font-size: 15px; color: #0f172a; margin-top: 0;">Dear <strong>${name}</strong>,</p>
        <p style="font-size: 14px; color: #334155; line-height: 1.5;">
          Congratulations! Your registration for <strong>The Asif Talks (Episode #3)</strong> has been confirmed. Below is your official boarding pass information:
        </p>
        
        <div style="background-color: #f8fafc; border: 2px dashed #cbd5e1; border-radius: 10px; padding: 18px; margin: 20px 0; text-align: center;">
          <div style="font-size: 11px; font-weight: bold; color: #64748b; text-transform: uppercase; letter-spacing: 1px;">Official Pass ID</div>
          <div style="font-size: 30px; font-weight: 800; color: #1d4ed8; margin: 6px 0; letter-spacing: 1px;">${passId}</div>
          <div style="font-size: 12px; color: #16a34a; font-weight: 700;">✓ Confirmed Live Audience Seat (Capacity: 300)</div>
        </div>
        
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 13px;">
          <tr>
            <td style="padding: 8px 0; color: #64748b; width: 35%; border-bottom: 1px solid #f1f5f9;"><strong>Distinguished Guest:</strong></td>
            <td style="padding: 8px 0; color: #0f172a; border-bottom: 1px solid #f1f5f9;"><strong>Hon. Smt. Meghana Sakore-Bordikar</strong><br><span style="font-size: 11.5px; color: #475569;">Minister of State, Govt. of Maharashtra & Guardian Minister, Parbhani District</span></td>
          </tr>
          <tr>
            <td style="padding: 8px 0; color: #64748b; border-bottom: 1px solid #f1f5f9;"><strong>Host & Creator:</strong></td>
            <td style="padding: 8px 0; color: #0f172a; border-bottom: 1px solid #f1f5f9;">Syed Asif (President, DCode Developers Club)</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; color: #64748b; border-bottom: 1px solid #f1f5f9;"><strong>Event Date & Time:</strong></td>
            <td style="padding: 8px 0; color: #b45309; font-weight: bold; border-bottom: 1px solid #f1f5f9;">Thursday, 1 October 2026 | 10:30 AM Sharp</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; color: #64748b;"><strong>Venue:</strong></td>
            <td style="padding: 8px 0; color: #0f172a;">Dr. A.P.J. Abdul Kalam Auditorium, SSIEMS, Parbhani</td>
          </tr>
        </table>
        
        <div style="background-color: #eff6ff; border-left: 4px solid #2563eb; padding: 14px; border-radius: 4px; margin-bottom: 22px;">
          <h4 style="margin: 0 0 8px 0; font-size: 13px; color: #1e40af;">📌 Important Event Guidelines:</h4>
          <ul style="margin: 0; padding-left: 18px; font-size: 12px; color: #1e3a8a; line-height: 1.6;">
            <li><strong>Mandatory Photo ID:</strong> Attendees must carry a physical College ID or Govt Photo ID card for entry verification.</li>
            <li><strong>Entry Time:</strong> Gate opens at <strong>10:30 AM Sharp</strong>. Auditorium doors will close once recording starts.</li>
            <li><strong>Media Policy:</strong> Photos and videos are allowed during the session (without original audio).</li>
            <li><strong>Verification:</strong> Present this confirmation email or your downloaded boarding pass QR code at the entrance desk.</li>
          </ul>
        </div>
        
        <div style="text-align: center; margin-top: 24px;">
          <a href="https://chat.whatsapp.com/EnpTg11Zxot9y1c8WLXr3s" style="background-color: #22c55e; color: #ffffff; padding: 12px 22px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 13px; display: inline-block;">
            Join Official Audience WhatsApp Group
          </a>
        </div>
      </div>
      
      <div style="background-color: #f8fafc; padding: 14px; text-align: center; font-size: 11px; color: #64748b; border-top: 1px solid #e2e8f0;">
        Organized by <strong>DCode Developers Club</strong> • In Partnership with <strong>SSIEMS Parbhani</strong> & <strong>Vertex Institute</strong><br>
        Official Website: <a href="https://www.dcode.club" style="color: #2563eb; text-decoration: none;">www.dcode.club</a>
      </div>
    </div>
  `;
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (err) {
    Logger.log("Email notification error: " + err.toString());
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevent concurrent write collisions

  try {
    var sheet = getEpisodeSheet();
    var lastRow = sheet.getLastRow();
    var p = e.parameter || {};

    // -------------------------------------------------------------
    // GATE SCANNER CHECK-IN ACTION
    // -------------------------------------------------------------
    if (p.action === "checkin") {
      var searchPass = (p.passId || "").trim().toUpperCase();
      var searchPhone = normalizePhone(p.phone || "");
      var volunteerName = (p.volunteer || "Gate Entry Volunteer").trim();

      if (!searchPass && !searchPhone) {
        return ContentService.createTextOutput(JSON.stringify({
          result: "error",
          message: "Pass ID or Mobile Number is required for check-in."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      if (lastRow <= 1) {
        return ContentService.createTextOutput(JSON.stringify({
          result: "not_found",
          message: "No registrations recorded in the database yet."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var maxCols = Math.max(14, sheet.getLastColumn());
      var data = sheet.getRange(2, 1, lastRow - 1, maxCols).getValues();
      var targetRow = -1;
      var matchedRecord = null;

      for (var i = 0; i < data.length; i++) {
        var rowPass = String(data[i][1] || "").trim().toUpperCase();
        var rowPhone = normalizePhone(data[i][3]);

        var isMatch = false;
        if (searchPass && rowPass === searchPass) isMatch = true;
        if (!isMatch && searchPhone && rowPhone === searchPhone) isMatch = true;

        if (isMatch) {
          targetRow = i + 2; // Row index (1-based, 2 is first data row)
          matchedRecord = {
            passId: rowPass,
            name: String(data[i][2] || "").trim(),
            phone: String(data[i][3] || "").trim(),
            email: String(data[i][4] || "").trim(),
            gender: String(data[i][5] || "").trim(),
            college: String(data[i][6] || "").trim(),
            city: String(data[i][7] || "").trim(),
            question: String(data[i][9] || "").trim(),
            currentStatus: String(data[i][11] || "Pending").trim(),
            checkInTime: String(data[i][12] || "").trim(),
            volunteer: String(data[i][13] || "").trim()
          };
          break;
        }
      }

      if (targetRow === -1) {
        return ContentService.createTextOutput(JSON.stringify({
          result: "not_found",
          message: "Pass ID (" + (searchPass || searchPhone) + ") not found in registered attendees list!"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Check if already checked in
      if (p.force !== "true" && matchedRecord.currentStatus === "Checked In") {
        return ContentService.createTextOutput(JSON.stringify({
          result: "already_checked_in",
          attendee: matchedRecord,
          message: "ALREADY CHECKED IN: " + matchedRecord.name + " (" + matchedRecord.passId + ") was already admitted at " + (matchedRecord.checkInTime || "earlier today") + "!"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Mark Checked In in sheet
      var nowStr = Utilities.formatDate(new Date(), "GMT+5:30", "dd MMM yyyy, hh:mm:ss a");
      sheet.getRange(targetRow, 12).setValue("Checked In").setBackground("#dcfce7").setFontColor("#15803d").setFontWeight("bold");
      sheet.getRange(targetRow, 13).setValue(nowStr);
      sheet.getRange(targetRow, 14).setValue(volunteerName);

      matchedRecord.currentStatus = "Checked In";
      matchedRecord.checkInTime = nowStr;
      matchedRecord.volunteer = volunteerName;

      return ContentService.createTextOutput(JSON.stringify({
        result: "success",
        attendee: matchedRecord,
        message: "Entry Confirmed! Welcome " + matchedRecord.name + " (" + matchedRecord.passId + ")."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------
    // REGULAR ATTENDEE REGISTRATION
    // -------------------------------------------------------------
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

    // Validate phone number: must be at least 10 digits
    if (!cleanPhone || cleanPhone.length < 10) {
      return ContentService
        .createTextOutput(JSON.stringify({ 
          result: "error",
          error: "INVALID_PHONE",
          message: "Please enter a valid mobile number with at least 10 digits." 
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Validate email address: required
    if (!email || email.indexOf("@") === -1) {
      return ContentService
        .createTextOutput(JSON.stringify({ 
          result: "error",
          error: "INVALID_EMAIL",
          message: "A valid email address is required to receive your pass and guidelines." 
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Check for duplicate mobile number in Column D (Rows 2 to lastRow)
    if (lastRow > 1) {
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
              message: "This mobile number is already registered for The Asif Talks (Episode #3)! Pass ID: " + existingPassId 
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
      consent,
      "Pending", // Check-in Status
      "",        // Check-in Time
      ""         // Gate Volunteer
    ];

    sheet.appendRow(rowData);

    // Send automatic Pass confirmation email with guidelines
    try {
      sendConfirmationEmail(email, name, passId);
    } catch (mailErr) {
      Logger.log("Mail send failure: " + mailErr.toString());
    }

    var totalSeats = 300;
    var remainingSeats = Math.max(0, totalSeats - passCount);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        result: "success", 
        passId: passId, 
        registeredCount: passCount,
        remainingSeats: remainingSeats,
        message: "Registration for Episode #3 recorded successfully! Pass sent to your email." 
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
  var totalSeats = 300;
  var remainingSeats = Math.max(0, totalSeats - count);
  var nextPassId = "TAT-" + ("000" + (count + 1)).slice(-3);

  // -------------------------------------------------------------
  // SCANNER SYNC / ATTENDEES LIST
  // -------------------------------------------------------------
  if (e && e.parameter && (e.parameter.action === "scanner_sync" || e.parameter.action === "get_attendees")) {
    var attendeeList = [];
    var checkedInCount = 0;

    if (lastRow > 1) {
      var maxCols = Math.max(14, sheet.getLastColumn());
      var allRows = sheet.getRange(2, 1, lastRow - 1, maxCols).getValues();

      for (var j = 0; j < allRows.length; j++) {
        var rPass = String(allRows[j][1] || "").trim();
        var rName = String(allRows[j][2] || "").trim();
        var rPhone = normalizePhone(allRows[j][3]);
        var rCollege = String(allRows[j][6] || "").trim();
        var rCity = String(allRows[j][7] || "").trim();
        var rQuestion = String(allRows[j][9] || "").trim();
        var rStatus = String(allRows[j][11] || "Pending").trim();
        var rTime = String(allRows[j][12] || "").trim();

        if (rStatus === "Checked In") checkedInCount++;

        attendeeList.push({
          passId: rPass,
          name: rName,
          phone: rPhone,
          college: rCollege,
          city: rCity,
          question: rQuestion,
          status: rStatus === "Checked In" ? "Checked In" : "Pending",
          checkInTime: rTime
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "ok",
      episode: "Episode #3",
      guest: "Hon. Smt. Meghana Sakore-Bordikar",
      totalSeats: totalSeats,
      registeredCount: attendeeList.length,
      checkedInCount: checkedInCount,
      pendingCount: Math.max(0, attendeeList.length - checkedInCount),
      attendees: attendeeList
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // -------------------------------------------------------------
  // DEFAULT REGISTRATION STATUS & DUPLICATE PHONE CHECK
  // -------------------------------------------------------------
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
    isClosed: false,
    episode: "Episode #3",
    guest: "Hon. Smt. Meghana Sakore-Bordikar",
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
  Logger.log("Episode 3 Sheet successfully connected: " + sheet.getName());
}
