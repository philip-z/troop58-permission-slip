function doGet(e) {
  try {
    // ROUTER: If frontend requests a profile, return it and stop.
    if (e.parameter.action === 'getProfile' && e.parameter.googleUserId) {
      var profile = getUserProfile(e.parameter.googleUserId);
      return ContentService.createTextOutput(JSON.stringify(profile || {}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Searches for 'Trip Config', 'Trips', or defaults to the first tab
    var sheet = ss.getSheetByName('Trips');
    if (!sheet) {
      throw new Error("Trips configuration Missing");
    }    

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var config = {};
    
    if (lastRow > 1 && lastCol > 0) {
      // 1. Read raw headers from Row 1 and raw trip data from the latest row
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var tripData = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];
      
      // 2. Map exact sheet column headers directly as keys
      headers.forEach(function(header, idx) {
        if (header && header.toString().trim() !== "") {
          var key = header.toString().trim();
          config[key] = tripData[idx];
        }
      });
    }

    return ContentService.createTextOutput(JSON.stringify(config))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    
    // ROUTER: If frontend is saving a profile, route it and stop.
    if (data.action === 'saveProfile' && data.googleUserId) {
      saveOrUpdateUserProfile(data.googleUserId, data.profileData || {});
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ARCHIVE: Generate a PDF of this submission and file it in Drive under the trip's folder.
    // A PDF failure must never block the submission itself, so it is caught and logged.
    // Always overwrite 'PDF Link' so a client can't spoof it.
    data['PDF Link'] = '';
    try {
      data['PDF Link'] = archiveSubmissionPdf_(data).getUrl();
    } catch (pdfErr) {
      console.error('PDF archive failed: ' + pdfErr);
    }

    // ... [Rest of your existing doPost code for saving Responses] ...
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Responses') || ss.getSheets()[1];
    
    var lastCol = sheet.getLastColumn();
    var headers = [];
    
    // 1. Read existing exact headers from Row 1
    if (sheet.getLastRow() > 0 && lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
        return h.toString().trim();
      });
    }
    
    // 2. Ensure "Timestamp" is in Column 1
    if (headers.length === 0 || headers[0] === "") {
      headers = ["Timestamp"];
      sheet.getRange(1, 1, 1, 1).setValues([headers]).setFontWeight("bold");
    }
    
    // 3. Map exact header strings to column array indexes
    var headerMap = {};
    headers.forEach(function(h, idx) {
      if (h !== "") {
        headerMap[h] = idx;
      }
    });
    
    // 4. Inspect incoming JSON keys (skipping internal metadata like _tripId)
    Object.keys(data).forEach(function(key) {
      if (key.startsWith('_')) return; // Ignore internal metadata
      
      // If the incoming JSON key is NOT an exact match to a column header, add it!
      if (!(key in headerMap)) {
        var newColIdx = headers.length;
        
        // Write the incoming key directly as the new column header title in Row 1
        sheet.getRange(1, newColIdx + 1).setValue(key).setFontWeight("bold");
        
        headers.push(key);
        headerMap[key] = newColIdx;
      }
    });
    
    // 5. Build row array matching header order exactly
    var rowData = headers.map(function(header) {
      if (header.toLowerCase() === "timestamp") return new Date();
      
      return (data[header] !== undefined && data[header] !== null) 
        ? data[header] 
        : "";
    });
    
    // 6. Append entry row
    sheet.appendRow(rowData);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------------------------------------------------------------------------
// PDF ARCHIVE
// Layout in Drive (next to the spreadsheet):
//   <folder containing the sheet>/Permission Slips/<Trip title - dates>/<slip>.pdf
// Duplicate submissions are all kept: every filename carries a timestamp.
// ---------------------------------------------------------------------------

function archiveSubmissionPdf_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var trip = getLatestTripConfig_(ss);

  // Only the folder lookup/creation is locked, so two simultaneous submissions
  // can't create duplicate trip folders.
  var tripFolder;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var root = DriveApp.getFileById(ss.getId()).getParents().next();
    var slipsFolder = getOrCreateFolder_(root, 'Permission Slips');
    tripFolder = getOrCreateFolder_(slipsFolder, tripFolderName_(trip));
  } finally {
    lock.releaseLock();
  }

  var pdf = HtmlService.createHtmlOutput(renderSlipHtml_(data, trip, now))
    .getBlob()
    .getAs('application/pdf')
    .setName(slipFileName_(data, now) + '.pdf');

  return tripFolder.createFile(pdf);
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// Reads the latest row of the Trips sheet (same source doGet uses). The trip
// always comes from the server, never from the browser.
function getLatestTripConfig_(ss) {
  var config = {};
  var sheet = ss.getSheetByName('Trips');
  if (!sheet) return config;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return config;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];
  headers.forEach(function(h, i) {
    var key = h.toString().trim();
    if (key !== '') config[key] = row[i];
  });
  return config;
}

function displayValue_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'MMM d, yyyy');
  }
  return v === undefined || v === null ? '' : String(v);
}

function safeName_(s) {
  return String(s || '')
    .replace(/[\\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 100);
}

function tripFolderName_(trip) {
  var title = displayValue_(trip.tripTitle);
  var dates = displayValue_(trip.tripDates);
  var name = safeName_(title + (dates ? ' - ' + dates : ''));
  return name || 'Unknown Trip';
}

function slipFileName_(data, now) {
  var isAdult = String(data['Form Type']).toLowerCase() === 'adult';
  var person = isAdult ? data['Adult Name'] : data['Youth Name'];
  var stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HHmmss');
  return safeName_((isAdult ? 'Adult' : 'Youth') + ' - ' + (person || 'Unnamed') + ' - ' + stamp);
}

function escapeHtml_(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// Builds a simple HTML document for the PDF. Apps Script's HTML-to-PDF
// converter supports only basic CSS, so this is a clean record of the slip,
// not a pixel copy of the web page. EVERY submitted value is escaped.
// NOTE: the statements below are a copy of the wording on the web form. If the
// wording in index.html changes, update it here too.
function renderSlipHtml_(data, trip, now) {
  var isAdult = String(data['Form Type']).toLowerCase() === 'adult';
  var tz = Session.getScriptTimeZone();

  var statements = isAdult
    ? ['I waive all claims against the leaders of this trip and their families and give permission for them to seek any emergency medical or surgical care deemed necessary for me, should the need arise.']
    : ['I waive all claims against the leaders of this trip and their families and give permission for them to seek any emergency medical or surgical care deemed necessary for my son, should the need arise.',
       'I have read and understand this form, and I agree to comply with Troop 58\u2019s electronic device policies to provide a distraction free outdoor experience for all participants.'];
  var esign = 'By typing their full legal name and clicking Submit, the signer acknowledged that the typed name is their official, legal, and binding signature, equivalent to a handwritten signature.';

  var skip = { 'Form Type': true, 'PDF Link': true, 'action': true };
  var rows = Object.keys(data).filter(function(k) {
    return !k.startsWith('_') && !skip[k];
  }).map(function(k) {
    var v = data[k];
    if (Array.isArray(v)) v = v.join(', ');
    var isSig = k.toLowerCase().indexOf('signature') !== -1 && k.toLowerCase().indexOf('date') === -1;
    var valHtml = escapeHtml_(v);
    if (isSig) valHtml = '<i style="font-size:15pt;">' + valHtml + '</i>';
    return '<tr><td style="width:38%;padding:4px 6px;border-bottom:1px solid #ddd;color:#444;"><b>' +
      escapeHtml_(k) + '</b></td><td style="padding:4px 6px;border-bottom:1px solid #ddd;">' + valHtml + '</td></tr>';
  }).join('');

  var tripLines = [
    ['Location', trip.tripLocation], ['Dates', trip.tripDates],
    ['Cost', trip.tripCost], ['Coordinator', trip.tripCoordinator]
  ].filter(function(p) { return displayValue_(p[1]) !== ''; })
   .map(function(p) { return '<b>' + p[0] + ':</b> ' + escapeHtml_(displayValue_(p[1])); })
   .join(' &nbsp;|&nbsp; ');

  return '<html><body style="font-family:Times New Roman,serif;font-size:11pt;">' +
    '<h2 style="text-align:center;margin:0;">BSA Troop 58 Elsmere</h2>' +
    '<h3 style="text-align:center;margin:2px 0;">' + escapeHtml_(displayValue_(trip.tripTitle)) + '</h3>' +
    '<p style="text-align:center;margin:2px 0;">' + tripLines + '</p>' +
    '<h3 style="text-align:center;margin:12px 0 6px;">' + (isAdult ? 'ADULT' : 'YOUTH') + ' PERMISSION SLIP</h3>' +
    '<table style="width:100%;border-collapse:collapse;">' + rows + '</table>' +
    '<p style="margin-top:14px;">' + statements.map(escapeHtml_).join('<br><br>') + '</p>' +
    '<p style="font-size:10pt;color:#333;">' + escapeHtml_(esign) + '</p>' +
    '<p style="font-size:9pt;color:#666;">Submitted electronically: ' +
      Utilities.formatDate(now, tz, 'MMM d, yyyy h:mm:ss a z') + '</p>' +
    '</body></html>';
}

// Fetch saved profile by Google User ID dynamically
function getUserProfile(googleUserId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Profiles');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null; // No users saved yet

    // 1. Read exact headers from Row 1
    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var id = String(googleUserId);
    
    // 2. Search row by row for matching Google ID in Column A (index 0).
    //    Compare as strings: Sheets may hand back the ID as a number.
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        
        var profile = {};
        // 3. Dynamically map every sheet column to the profile object
        for (var j = 1; j < headers.length; j++) {
          var key = headers[j];
          if (key === "") continue;
          var v = data[i][j];
          // Dates must be yyyy-MM-dd for <input type="date"> to accept them
          if (v instanceof Date) {
            v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
          profile[key] = v;
        }
        return profile;
      }
    }
    return null; // Profile not found
  } catch (err) {
    console.error("Error fetching profile: " + err);
    return null;
  }
}

// Save or update profile on submission dynamically
function saveOrUpdateUserProfile(googleUserId, profileData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Profiles');
  var headers = [];
  var id = String(googleUserId);
  
  // 1. Ensure the Profiles sheet and the ID column exist
  if (!sheet) {
    sheet = ss.insertSheet('Profiles');
    headers = ["Google ID"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1).setFontWeight("bold");
  } else {
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
        return h.toString().trim();
      });
    }
    if (headers.length === 0 || headers[0] === "") {
      headers = ["Google ID"];
      sheet.getRange(1, 1).setValue("Google ID").setFontWeight("bold");
    }
  }

  var headerMap = {};
  headers.forEach(function(h, idx) {
    if (h !== "") headerMap[h] = idx;
  });

  // 2. Check incoming profileData for NEW fields and create columns automatically
  Object.keys(profileData).forEach(function(key) {
    if (key.startsWith('_') || key === 'Google ID') return; // Skip metadata
    
    if (!(key in headerMap)) {
      var newColIdx = headers.length;
      sheet.getRange(1, newColIdx + 1).setValue(key).setFontWeight("bold");
      headers.push(key);
      headerMap[key] = newColIdx;
    }
  });

  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;

  // 3. Find if this user already has a row (string comparison)
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      rowIndex = i + 1; // Apps Script ranges are 1-based, array is 0-based
      break;
    }
  }

  var existingRow = rowIndex > 0 ? data[rowIndex - 1] : [];

  // 4. Build the row payload. Fields not in this submission (e.g. the other
  //    form type's section, which is disabled) keep their previously saved value.
  var rowPayload = headers.map(function(header, idx) {
    if (header === "Google ID") return id;

    var incoming = profileData[header];
    if (incoming !== undefined && incoming !== null) {
      return Array.isArray(incoming) ? incoming.join(', ') : incoming;
    }
    var prev = existingRow[idx];
    return prev !== undefined ? prev : "";
  });

  // 5. Update the existing row, or write a new one. Force column A to plain
  //    text so the long numeric Google ID is not converted to a rounded number.
  var targetRow = rowIndex > 0 ? rowIndex : sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 1).setNumberFormat('@');
  sheet.getRange(targetRow, 1, 1, rowPayload.length).setValues([rowPayload]);
}
