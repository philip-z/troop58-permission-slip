// OAuth client ID that Google Sign-In tokens must be issued for. For production,
// set a Script Property named OAUTH_CLIENT_ID instead of editing this default.
var DEFAULT_OAUTH_CLIENT_ID = '1059045077738-ohkb0nupq7kh69kj3i5io9n1rbjmfn8m.apps.googleusercontent.com';

// Per-request abuse limits. These only FLAG or trim; they never reject a submission,
// and there is deliberately NO global rate limit (a burst of real parents must not be dropped).
var MAX_FIELD_LENGTH = 1000;
var MAX_KEY_LENGTH = 80;
var MAX_FIELDS = 80;
var MAX_PAYLOAD_CHARS = 60000;
var HONEYPOT_FIELD = 'website_url';

function doGet(e) {
  try {
    // Public, non-sensitive trip configuration only. Profiles are no longer served via GET.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Trips');
    if (!sheet) {
      throw new Error("Trips configuration Missing");
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var config = {};

    if (lastRow > 1 && lastCol > 0) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var tripData = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];

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

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var raw = e.postData.contents;
    var data = JSON.parse(raw);

    // PROFILE ROUTES: the user ID comes ONLY from a verified Google ID token.
    // Any googleUserId sent by the client is ignored.
    if (data.action === 'getProfile') {
      var readerId = verifyGoogleToken_(data.idToken);
      if (!readerId) return jsonOut_({ error: 'unauthorized' });
      return jsonOut_(getUserProfile(readerId) || {});
    }

    if (data.action === 'saveProfile') {
      var writerId = verifyGoogleToken_(data.idToken);
      if (!writerId) return jsonOut_({ status: 'error', message: 'unauthorized' });
      saveOrUpdateUserProfile(writerId, sanitizeFields_(data.profileData || {}, []));
      return jsonOut_({ status: 'success' });
    }

    // SUBMISSION (anonymous, no sign-in required).
    // Soft anti-spam: trim and flag, never reject.
    var flags = [];
    if (raw.length > MAX_PAYLOAD_CHARS) flags.push('Oversized payload');
    var honeypotFilled = !!(data[HONEYPOT_FIELD] && String(data[HONEYPOT_FIELD]).trim() !== '');
    if (honeypotFilled) flags.push('Honeypot filled');

    data = sanitizeFields_(data, flags);

    var isAdult = String(data['Form Type']).toLowerCase() === 'adult';
    if (!(isAdult ? data['Adult Name'] : data['Youth Name'])) flags.push('Missing name');
    if (isAdult ? !data['Adult Signature'] : (!data['Parent Signature'] || !data['Youth Signature'])) flags.push('Missing signature');

    // Always overwrite server-controlled columns so a client can't spoof them.
    data['PDF Link'] = '';
    // A honeypot hit is still saved as a row for review, but doesn't get a PDF in Drive.
    // A PDF failure must never block the submission itself.
    if (!honeypotFilled) {
      try {
        data['PDF Link'] = archiveSubmissionPdf_(data).getUrl();
      } catch (pdfErr) {
        console.error('PDF archive failed: ' + pdfErr);
        flags.push('PDF failed');
      }
    }
    data['Flags'] = flags.join('; ');

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

    // 4. Add a column for any new incoming key
    Object.keys(data).forEach(function(key) {
      if (!(key in headerMap)) {
        var newColIdx = headers.length;
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

    return jsonOut_({ status: "success" });

  } catch (err) {
    return jsonOut_({ status: "error", message: err.toString() });
  }
}

// ---------------------------------------------------------------------------
// AUTH + INPUT HYGIENE
// ---------------------------------------------------------------------------

function getOAuthClientId_() {
  return PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID') || DEFAULT_OAUTH_CLIENT_ID;
}

// Returns the verified Google user ID (the token's 'sub'), or null if the token is invalid.
function verifyGoogleToken_(idToken) {
  if (!idToken || typeof idToken !== 'string') return null;
  try {
    var cache = CacheService.getScriptCache();
    var key = 'tok_' + Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
    var hit = cache.get(key);
    if (hit) return hit;

    var res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;

    var info = JSON.parse(res.getContentText());
    if (info.aud !== getOAuthClientId_()) return null;      // must be minted for THIS client
    if (Number(info.exp) * 1000 < Date.now()) return null;  // not expired
    if (!info.sub) return null;

    var ttl = Math.min(300, Math.floor(Number(info.exp) - Date.now() / 1000));
    if (ttl > 0) cache.put(key, String(info.sub), ttl);
    return String(info.sub);
  } catch (err) {
    console.error('Token verification failed: ' + err);
    return null;
  }
}

// Returns a cleaned copy: drops internal/server-controlled keys, caps key count and
// value length (flagging when it trims), and flattens arrays/objects to strings.
function sanitizeFields_(input, flags) {
  var blocked = { 'action': true, 'idToken': true, 'googleUserId': true,
                  'Flags': true, 'PDF Link': true };
  blocked[HONEYPOT_FIELD] = true;

  var clean = {};
  var count = 0;
  Object.keys(input || {}).forEach(function(key) {
    if (key.charAt(0) === '_' || blocked[key] || key.length > MAX_KEY_LENGTH) return;
    if (count >= MAX_FIELDS) {
      if (flags.indexOf('Too many fields') === -1) flags.push('Too many fields');
      return;
    }
    var v = input[key];
    if (v === undefined || v === null) v = '';
    if (Array.isArray(v)) v = v.join(', ');
    else if (typeof v === 'object') v = JSON.stringify(v);
    else v = String(v);
    if (v.length > MAX_FIELD_LENGTH) {
      v = v.substring(0, MAX_FIELD_LENGTH);
      if (flags.indexOf('Field trimmed') === -1) flags.push('Field trimmed');
    }
    clean[key] = v;
    count++;
  });
  return clean;
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

  var skip = { 'Form Type': true, 'PDF Link': true, 'Flags': true, 'action': true };
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

// Fetch saved profile by (verified) Google User ID
function getUserProfile(googleUserId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Profiles');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null; // No users saved yet

    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var id = String(googleUserId);

    // Compare as strings: Sheets may hand back the ID as a number.
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {

        var profile = {};
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

// Save or update profile for a (verified) Google User ID
function saveOrUpdateUserProfile(googleUserId, profileData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
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

    // 2. Create columns for new profile fields
    Object.keys(profileData).forEach(function(key) {
      if (key.startsWith('_') || key === 'Google ID') return;

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

    // 4. Fields not in this save keep their previously saved value.
    var rowPayload = headers.map(function(header, idx) {
      if (header === "Google ID") return id;

      var incoming = profileData[header];
      if (incoming !== undefined && incoming !== null) {
        return Array.isArray(incoming) ? incoming.join(', ') : incoming;
      }
      var prev = existingRow[idx];
      return prev !== undefined ? prev : "";
    });

    // 5. Force column A to plain text so the long numeric Google ID isn't rounded.
    var targetRow = rowIndex > 0 ? rowIndex : sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setNumberFormat('@');
    sheet.getRange(targetRow, 1, 1, rowPayload.length).setValues([rowPayload]);
  } finally {
    lock.releaseLock();
  }
}
