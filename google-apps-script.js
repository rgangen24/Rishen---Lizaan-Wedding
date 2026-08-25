/**
 * =========================================================================
 * RISHEN & LIZAAN WEDDING INVITATION — GOOGLE APPS SCRIPT WEB BACKEND
 * =========================================================================
 * 
 * RECIPIENT EMAILS:
 * - rgangen24@gmail.com / rgangen@gmail.com
 * - lizaan.tait2@gmail.com
 * =========================================================================
 */

// NOTIFICATION EMAILS (Both Rishen & Lizaan will receive alerts)
var NOTIFICATION_EMAILS = "rgangen24@gmail.com, rgangen@gmail.com, lizaan.tait2@gmail.com";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var raw = e.postData ? e.postData.contents : "{}";
    var data = {};
    
    try {
      data = JSON.parse(raw);
    } catch (err) {
      data = e.parameter || {};
    }

    var now = new Date();
    var formattedDate = Utilities.formatDate(now, "Africa/Johannesburg", "dd/MM/yyyy HH:mm:ss");
    var submissionId = Utilities.getUuid();

    var primaryName = data.primaryName || data.name || data.guestFullNameField || "Guest";
    var attendanceRaw = data.attendance || data.attendanceStatusRadio || "joyfully";
    var isAttending = (attendanceRaw === "joyfully" || attendanceRaw === "yes" || attendanceRaw.indexOf("Yes") !== -1);
    var attendanceStatus = isAttending ? "Yes, I'll be there" : "No, unfortunately I can't attend";

    var guestCount = isAttending ? (parseInt(data.guestCount || data.guestQuantitySelect, 10) || 1) : 0;
    var message = data.message || data.guestWishesTextarea || "";
    var generalDietary = data.dietary || data.guestDietaryNotesInput || "";

    // Parse guest list (Up to 5 guests)
    var guests = [];
    if (data.guests && Array.isArray(data.guests)) {
      guests = data.guests;
    } else {
      guests.push({
        name: primaryName,
        dietary: generalDietary || "No special requirements",
        dietaryDetail: ""
      });
      for (var i = 2; i <= guestCount; i++) {
        var gName = data["guest_" + i + "_name"] || ("Guest " + i);
        var gDiet = data["guest_" + i + "_dietary"] || "No special requirements";
        guests.push({
          name: gName,
          dietary: gDiet,
          dietaryDetail: ""
        });
      }
    }

    // Build Row matching exact Google Sheet schema:
    // Timestamp | Submission ID | Primary guest name | Attendance | Number of guests | Guest 1 name | Guest 1 dietary | Guest 1 dietary detail | Guest 2... | Guest 3... | Guest 4... | Guest 5... | Message
    var row = [
      formattedDate,
      submissionId,
      primaryName,
      attendanceStatus,
      guestCount.toString()
    ];

    for (var g = 0; g < 5; g++) {
      if (g < guests.length && isAttending) {
        row.push(guests[g].name || "");
        row.push(guests[g].dietary || "No special requirements");
        row.push(guests[g].dietaryDetail || "");
      } else {
        row.push("");
        row.push("");
        row.push("");
      }
    }

    row.push(message);

    sheet.appendRow(row);

    // Send Instant Email Notification to both Rishen & Lizaan
    sendEmailNotification(primaryName, attendanceStatus, guestCount, guests, message, formattedDate);

    return ContentService
      .createTextOutput(JSON.stringify({
        result: "success",
        submissionId: submissionId,
        message: "RSVP recorded successfully."
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        result: "error",
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];
    var data = [];

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var entry = {};
      for (var j = 0; j < headers.length; j++) {
        entry[headers[j]] = row[j];
      }
      data.push(entry);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        result: "success",
        total: data.length,
        data: data
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        result: "error",
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendEmailNotification(primaryName, attendanceStatus, guestCount, guests, message, dateStr) {
  var isAttending = (attendanceStatus.indexOf("Yes") !== -1);
  var subject = (isAttending ? "💍 RSVP ACCEPTED: " : "💌 RSVP DECLINED: ") + primaryName;

  var guestListHtml = "";
  if (isAttending && guests.length > 0) {
    guestListHtml = "<h3 style='color:#4D161D; margin-top:16px; margin-bottom:8px;'>Guest Details:</h3><ul style='margin:0; padding-left:20px;'>";
    for (var i = 0; i < guests.length; i++) {
      guestListHtml += "<li style='margin-bottom:4px;'><strong>" + (guests[i].name || ("Guest " + (i + 1))) + "</strong> &mdash; Dietary: " + (guests[i].dietary || "None") + "</li>";
    }
    guestListHtml += "</ul>";
  }

  var htmlBody = ""
    + "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #E7D9BE; border-radius: 8px; overflow: hidden;'>"
    + "  <div style='background: #4D161D; color: #FDFBF7; padding: 24px; text-align: center;'>"
    + "    <h1 style='margin: 0; font-size: 24px; font-weight: 300;'>Rishen &amp; Lizaan Wedding</h1>"
    + "    <p style='margin: 8px 0 0; color: #E7D9BE; font-size: 14px;'>New RSVP Submission Received</p>"
    + "  </div>"
    + "  <div style='padding: 24px; background: #FAF7F2; color: #26201B;'>"
    + "    <p style='font-size: 18px; margin-top: 0;'><strong>Guest:</strong> " + primaryName + "</p>"
    + "    <p style='font-size: 16px;'><strong>Status:</strong> <span style='color: " + (isAttending ? "#2E7D32" : "#C62828") + "; font-weight: bold;'>" + attendanceStatus + "</span></p>"
    + "    <p style='font-size: 16px;'><strong>Total Guests:</strong> " + guestCount + "</p>"
    + guestListHtml
    + (message ? "<div style='margin-top: 20px; padding: 15px; background: #FFFFFF; border-left: 4px solid #6E2A31; border-radius: 4px;'><p style='margin:0; font-style:italic;'>&ldquo;" + message + "&rdquo;</p></div>" : "")
    + "    <p style='font-size: 12px; color: #857A70; margin-top: 24px;'>Submitted on " + dateStr + " &middot; View live sheet <a href='https://docs.google.com/spreadsheets/d/1b0m89-RZq79QvfhU3DkczC4Y262kbVvcO9WUVlQcCL4/edit' style='color: #6E2A31;'>here</a></p>"
    + "  </div>"
    + "</div>";

  try {
    MailApp.sendEmail({
      to: NOTIFICATION_EMAILS,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (e) {
    Logger.log("Email error: " + e);
  }
}
