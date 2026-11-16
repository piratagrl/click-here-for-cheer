/**
 * Holiday Gingerbread Decorating Contest
 * Server Backend (Code.gs)
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('sleighride')
    .setTitle('Holiday Gingerbread Decorating Contest')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎁 Cheer Logs')
    .addItem('Launch Merriment', 'deployCheer')
    .addToUi();
}

function showWebUrl() {
  deployCheer();
}

function deployCheer() {
  var webAppUrl = "https://script.google.com/a/macros/google.com/s/AKfycbx2ECAFUZQQIGNeEWEATnPaxbGgkEx9fAty82O4izLCyTmJyLJvAAnxHXfTslpNxGi4LA/exec"; 

  var html = HtmlService.createHtmlOutput(
    "<div style='text-align:center; padding-top:15px; font-family:sans-serif;'>" +
    "<p style='color:#555; font-size:14px; margin-bottom:20px;'>Your secure workstation settings require a dedicated window layer.</p>" +
    
    // NATIVE HTML FORM BUTTON CONTAINER
    "<form action='" + webAppUrl + "' target='_blank'>" +
    "<button type='submit' style='display:inline-block; width:85%; padding:14px; background:#b22222; color:white; border:none; border-radius:6px; font-weight:bold; font-size:16px; cursor:pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>🎁 Click for Holiday Spirit 🎁</button>" +
    "</form>" +
    
    "</div>"
  ).setWidth(400).setHeight(160);

  SpreadsheetApp.getUi().showModalDialog(html, 'Connection Security Initialized');
}

/**
 * Saves decorated house submission to its own dedicated tab & Google Drive.
 */
function saveParagraphToSheet(dataUrl, title, baker) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contest Submissions') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];

  // 1. Save full resolution snapshot to Drive
  var folders = DriveApp.getFoldersByName("Gingerbread Contest Submissions");
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Gingerbread Contest Submissions");
  var cleanBaker = (baker || "Baker").replace(/[^a-zA-Z0-9]/g, '_');
  var cleanTitle = (title || "House").replace(/[^a-zA-Z0-9]/g, '_');
  var fileName = cleanBaker + "_" + cleanTitle + "_" + new Date().getTime() + ".jpg";
  var bytes = Utilities.base64Decode(dataUrl.split(',')[1]);
  var file = folder.createFile(fileName, bytes, "image/jpeg");

  // 2. Automatically generate the Column F thumbnail
  var token = ScriptApp.getOAuthToken();
  var thumbBase64 = "";
  try {
    var response = UrlFetchApp.fetch("https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w220", {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      thumbBase64 = "data:image/jpeg;base64," + Utilities.base64Encode(response.getBlob().getBytes());
    }
  } catch(e) {}

  // 3. Safety Guard: Downscale to 160px if house has every decoration and exceeds 45k chars
  if (thumbBase64.length > 45000) {
    try {
      var fbRes = UrlFetchApp.fetch("https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w160", {
        headers: { Authorization: "Bearer " + token },
        muteHttpExceptions: true
      });
      if (fbRes.getResponseCode() === 200) {
        thumbBase64 = "data:image/jpeg;base64," + Utilities.base64Encode(fbRes.getBlob().getBytes());
      }
    } catch(err) {}
  }

  var email = Session.getActiveUser().getEmail() || "";

  // Append new row with thumbnail in Column F:
  // Col A: Timestamp | Col B: Baker | Col C: Email | Col D: Title | Col E: Drive Link | Col F: Thumbnail | Col G: Hearts | Col H: Likes | Col I: Smileys
  sheet.appendRow([new Date(), baker, email, title, file.getUrl(), thumbBase64, 0, 0, 0]);
  return "House snapshot saved successfully!";
}

/**
 * Test function to authorize DriveApp and verify images.
 * Select 'testGallery' in the top toolbar and click 'Run'.
 */
function testGallery() {
  var entries = getGallerySubmissions();
  Logger.log("=== GALLERY TEST RESULTS ===");
  Logger.log("Total entries found: " + entries.length);
  for (var i = 0; i < entries.length; i++) {
    var imgLen = entries[i].image ? entries[i].image.length : 0;
    Logger.log((i + 1) + ". " + entries[i].title + " by " + entries[i].baker + " | Image Data Length: " + imgLen + " characters");
  }
}

/**
 * Extracts Google Drive File ID from links or raw IDs
 */
function extractDriveId(cellVal) {
  if (!cellVal) return '';
  var str = String(cellVal).trim();
  // 1. Matches /file/d/FILE_ID/ or /d/FILE_ID/
  var match = str.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (match && match[1]) return match[1];
  // 2. Matches ?id=FILE_ID
  match = str.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
  if (match && match[1]) return match[1];
  // 3. Matches raw 25+ char ID
  match = str.match(/[a-zA-Z0-9_-]{25,}/);
  return match ? match[0] : '';
}
/**
 * 1. RUN THIS to verify all images load and see their sizes in the log!
 */
function debugAndShowImages() {
  Logger.log("=== CHECKING GINGERBREAD SUBMISSIONS ===");
  
  var folders = DriveApp.getFoldersByName("Gingerbread Contest Submissions");
  if (!folders.hasNext()) {
    Logger.log("⚠️ Folder 'Gingerbread Contest Submissions' not found!");
    return;
  }
  
  var folder = folders.next();
  Logger.log("✅ Found folder: " + folder.getName());
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    Logger.log("   📸 File: " + f.getName() + " | Size: " + f.getSize() + " bytes");
  }
  
  var entries = getGallerySubmissions();
  Logger.log("=== RETRIEVED " + entries.length + " GALLERY ENTRIES ===");
  for (var i = 0; i < entries.length; i++) {
    Logger.log((i + 1) + ". " + entries[i].title + " by " + entries[i].baker + " | Image bytes: " + (entries[i].image ? entries[i].image.length : 0));
  }
}

/**
 * Reads submissions and matches photos from 'Gingerbread Contest Submissions'
 */
function getGallerySubmissions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contest Submissions') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 9)).getValues();
    var entries = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var baker = row[1] ? String(row[1]).trim() : 'Holiday Baker';
      var title = row[3] ? String(row[3]).trim() : '';
      var photoThumbnail = row[5] ? String(row[5]).trim() : ''; // Column F!
      var hearts = Number(row[6]) || 0;
      var likes = Number(row[7]) || 0;
      var smileys = Number(row[8]) || 0;
      var score = (hearts * 6) + (likes * 3) + (smileys * 1);

      if (!title || title.toLowerCase() === 'house title') continue;

      entries.push({
        rowIndex: i + 2,
        sheetName: sheet.getName(),
        baker: baker,
        title: title,
        image: photoThumbnail,     // Column F thumbnail
        imageData: photoThumbnail, // Dual compatibility
        hearts: hearts,
        likes: likes,
        smileys: smileys,
        score: score
      });
    }

    entries.sort(function(a, b) { return b.score - a.score; });
    return entries;
  } catch(err) {
    return [];
  }
}

/**
 * Records votes (Hearts, Likes, Smileys) directly to Columns G, H, I.
 */
function recordVote(sheetName, rowIndex, voteType) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    if (!sheet) return false;
    
    if (!sheet.getRange(1, 7).getValue()) sheet.getRange(1, 7).setValue('Hearts');
    if (!sheet.getRange(1, 8).getValue()) sheet.getRange(1, 8).setValue('Likes');
    if (!sheet.getRange(1, 9).getValue()) sheet.getRange(1, 9).setValue('Smileys');
    
    var col = (voteType === 'heart') ? 7 : (voteType === 'like' ? 8 : 9);
    var cell = sheet.getRange(rowIndex, col);
    var currentVotes = Number(cell.getValue()) || 0;
    cell.setValue(currentVotes + 1);
    
    return true;
  } catch (e) {
    Logger.log("Voting error: " + e.toString());
    return false;
  }
}
function testSubmissions() {
  var list = getGallerySubmissions();
  Logger.log("TOTAL SUBMISSIONS LOADED: " + list.length);
  list.forEach(function(item, idx) {
    Logger.log("Entry " + (idx + 1) + ": " + item.title + " by " + item.baker + " | Image Bytes: " + (item.image ? item.image.length : 0));
  });
}

function checkDriveFiles() {
  var folders = DriveApp.getFoldersByName("Gingerbread Contest Submissions");
  if (!folders.hasNext()) {
    Logger.log("❌ Folder 'Gingerbread Contest Submissions' NOT found!");
    return;
  }
  var files = folders.next().getFiles();
  while (files.hasNext()) {
    var f = files.next();
    Logger.log("📁 File: " + f.getName() + " | " + f.getSize() + " bytes");
  }
}
function saveThumbnailsToSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contest Submissions') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var token = ScriptApp.getOAuthToken();
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  for (var i = 0; i < data.length; i++) {
    var link = String(data[i][4]); // Column E (Drive Link)
    var match = link.match(/[-\w]{25,}/);
    if (!match) continue;

    var fileId = match[0];
    try {
      // Fetch a clean 200px thumbnail from Google Drive using Apps Script authorization
      var response = UrlFetchApp.fetch("https://drive.google.com/thumbnail?id=" + fileId + "&sz=w220", {
        headers: { Authorization: "Bearer " + token },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        var base64 = "data:image/jpeg;base64," + Utilities.base64Encode(response.getBlob().getBytes());
        // Write the thumbnail directly into Column F (Row i + 2, Column 6)
        sheet.getRange(i + 2, 6).setValue(base64);
        Logger.log("✅ Row " + (i + 2) + " (" + data[i][3] + "): Thumbnail saved into Column F! (" + base64.length + " chars)");
      } else {
        Logger.log("⚠️ Row " + (i + 2) + " failed: Response code " + response.getResponseCode());
      }
    } catch(err) {
      Logger.log("❌ Error on row " + (i + 2) + ": " + err.message);
    }
  }
}
/**
 * Saves live community votes to Columns G, H, I in the Google Sheet.
 */
function recordVote(arg1, arg2, arg3) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Supports both recordVote(sheetName, rowIndex, type) and recordVote(rowIndex, type)
    var sheetName = (typeof arg1 === 'string' && isNaN(Number(arg1))) ? arg1 : null;
    var rowIndex = sheetName ? Number(arg2) : Number(arg1);
    var voteType = (sheetName ? arg3 : arg2) || 'like';

    var sheet = (sheetName && ss.getSheetByName(sheetName)) 
      || ss.getSheetByName('Contest Submissions') 
      || ss.getSheetByName('Sheet1') 
      || ss.getSheets()[0];

    if (!sheet || rowIndex < 2) return false;

    // Col G (7) = Hearts (+6) | Col H (8) = Likes (+3) | Col I (9) = Smileys (+1)
    var col = 8; // Default Likes
    var type = String(voteType).toLowerCase();
    if (type.indexOf('heart') !== -1 || type.indexOf('love') !== -1) {
      col = 7;
    } else if (type.indexOf('like') !== -1) {
      col = 8;
    } else if (type.indexOf('smiley') !== -1 || type.indexOf('smile') !== -1 || type.indexOf('nice') !== -1) {
      col = 9;
    }

    var cell = sheet.getRange(rowIndex, col);
    var currentCount = Number(cell.getValue()) || 0;
    cell.setValue(currentCount + 1);
    return true;
  } catch(err) {
    Logger.log("recordVote error: " + err.message);
    return false;
  }
}
