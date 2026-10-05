/**
 * ElectroNova 2026 — Registration backend
 * ----------------------------------------
 * WHAT THIS DOES
 * 1. Receives registration data POSTed from the website form
 * 2. Saves each registration as a new row in your Google Sheet:
 *    https://docs.google.com/spreadsheets/d/1IBlmilOg0s82lvb65RiF3rt-TnKOBoN3NU3mpG3JPuI/edit
 *    (open it any time -> File -> Download -> Microsoft Excel to get an .xlsx)
 * 3. Saves the abstract PDF into a Google Drive folder and links it in the sheet
 * 4. Emails a confirmation to the registrant's own address
 * 5. Emails a notification (with the same details) to the organiser inbox
 *
 * SETUP (one time, ~3 minutes — this version does NOT need to be opened
 * from inside the sheet; it's a standalone script that writes to your sheet
 * by its ID, which is already filled in below)
 * 1. Go to https://script.new — this opens a blank Apps Script project.
 * 2. Delete anything in the editor and paste this whole file in.
 * 3. Change ORGANISER_EMAIL below to the real inbox that should get notified.
 * 4. Click Deploy -> New deployment -> select type "Web app".
 *      - Description: anything
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy. The first time, Google will ask you to authorise the
 *    script (it needs permission to write to your sheet/drive and send
 *    email on your behalf) — click through "Advanced" -> "Go to project
 *    (unsafe)" if it warns you; this is your own script, so it's safe.
 * 5. Copy the "Web app URL" it gives you (ends in /exec).
 * 6. Paste that URL into the website's HTML where it says
 *    GOOGLE_SCRIPT_URL = "" — put the URL between the quotes.
 * 7. Re-publish the site. Done — registrations will now save into your
 *    sheet and email themselves automatically.
 *
 * If you ever change the code here, you must create a NEW deployment
 * (Deploy -> Manage deployments -> Edit -> New version) for the change
 * to take effect on the live site.
 */

var SHEET_ID = "1IBlmilOg0s82lvb65RiF3rt-TnKOBoN3NU3mpG3JPuI"; // your registrations sheet
var ORGANISER_EMAIL = "e18ecl301@egspec.org"; // <-- change if needed
var EVENT_NAME = "ElectroNova 2026";
var SHEET_NAME = "Registrations";
var DRIVE_FOLDER_NAME = "ElectroNova 2026 Abstracts";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var sheet = getOrCreateSheet_();
    var fileUrl = "";

    if (data.abstractBase64 && data.abstractFileName) {
      fileUrl = saveAbstractToDrive_(data.abstractBase64, data.abstractFileName, data.team);
    }

    sheet.appendRow([
      new Date(),
      data.team || "",
      data.size || "",
      data.lead || "",
      data.phone || "",
      data.email || "",
      data.college || "",
      data.track || "",
      data.year || "",
      data.abstractFileName || "",
      fileUrl
    ]);

    sendConfirmationEmail_(data);
    sendOrganiserEmail_(data, fileUrl);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var headerCell = sheet.getRange(1, 1).getValue();
  if (!headerCell) {
    sheet.appendRow([
      "Timestamp", "Team Name", "Team size", "Team lead's name", "Phone",
      "Email", "College", "Track", "Year of study", "Abstract file name", "Abstract link"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveAbstractToDrive_(base64Data, fileName, teamName) {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);

  var cleanBase64 = base64Data.indexOf(",") > -1 ? base64Data.split(",")[1] : base64Data;
  var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), "application/pdf", fileName);
  var safeTeam = (teamName || "team").replace(/[^a-zA-Z0-9 _-]/g, "");
  blob.setName(safeTeam + " — " + fileName);

  var file = folder.createFile(blob);
  return file.getUrl();
}

function sendConfirmationEmail_(data) {
  if (!data.email) return;

  var subject = "Confirmation of Registration – " + EVENT_NAME + " | " + data.team;
  var body =
    "Dear " + (data.lead || "Team Lead") + ",\n\n" +
    "Thank you for registering for " + EVENT_NAME + "! We have successfully received your submission.\n\n" +
    "Here is a summary of your registration details:\n\n" +
    "Registration Summary\n" +
    "Team Name: " + data.team + "\n" +
    "Team Size: " + data.size + "\n" +
    "College / Institution: " + data.college + "\n" +
    "Track: " + data.track + "\n" +
    "Year of Study: " + data.year + "\n" +
    "Abstract File: " + (data.abstractFileName || "(not attached)") + "\n\n" +
    "Keep this email as your confirmation. Shortlisted teams will be contacted " +
    "at this address with further details.\n\n" +
    "— " + EVENT_NAME + " Team\n" +
    "Dept. of Electronics & Communication Engineering, E.G.S. Pillay Engineering College";

  MailApp.sendEmail({ to: data.email, subject: subject, body: body });
}

function sendOrganiserEmail_(data, fileUrl) {
  var subject = "New registration — " + data.team;
  var body =
    "Team name: " + data.team + "\n" +
    "Team size: " + data.size + "\n" +
    "Team lead: " + data.lead + "\n" +
    "Phone: " + data.phone + "\n" +
    "Email: " + data.email + "\n" +
    "College: " + data.college + "\n" +
    "Track: " + data.track + "\n" +
    "Year: " + data.year + "\n" +
    (fileUrl ? "Abstract: " + fileUrl + "\n" : "Abstract: (not attached)\n");

  MailApp.sendEmail({ to: ORGANISER_EMAIL, subject: subject, body: body });
}
