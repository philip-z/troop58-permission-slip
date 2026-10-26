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

// Version label of the legal wording on the form (index.html). Bump it (e.g. a date) whenever
// the waiver, e-signature or other agreement text changes, and tag that commit in git, so every
// response row can be matched to the exact text it was signed under.
var FORM_VERSION = '2026-10-26';

// Diagnostic log: appends a row to the log sheet. Logging is toggled by renaming
// the sheet: 'Log(Enabled)' logs, 'Log(Disabled)' (or no sheet) does not.
// Never throws and never logs tokens or personal data.
var LOG_SHEET_ENABLED = 'Log(Enabled)';
var tokenFailReason_ = '';
function logEvent_(stage, detail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(LOG_SHEET_ENABLED);
    if (!sheet) return;
    sheet.appendRow([new Date(), stage, String(detail === undefined ? '' : detail).substring(0, 2000)]);
  } catch (e) {
    console.error('Log failed: ' + e);
  }
}

// Run this once from the editor (select it in the function dropdown, click Run).
// It exercises UrlFetch, Cache, Lock and the Log sheet so Google shows
// the permission prompt if any scope is missing, and records each result in Log.
function runDiagnostics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(LOG_SHEET_ENABLED)) {
    var disabled = ss.getSheetByName('Log(Disabled)');
    if (disabled) {
      disabled.setName(LOG_SHEET_ENABLED);
    } else {
      var created = ss.insertSheet(LOG_SHEET_ENABLED);
      created.appendRow(['Timestamp', 'Stage', 'Detail']);
      created.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
  }
  logEvent_('diag start', 'diagnostics started');
  var steps = {
    'UrlFetch': function() { return UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=x', { muteHttpExceptions: true }).getResponseCode(); },
    'Cache': function() { CacheService.getScriptCache().put('diag', '1', 5); return 'ok'; },
    'Lock': function() { var l = LockService.getScriptLock(); l.waitLock(5000); l.releaseLock(); return 'ok'; },
    'Profiles sheet': function() { return ss.getSheetByName('Profiles') ? 'exists' : 'missing (created on first save)'; },
    'Trips sheet': function() { return ss.getSheetByName('Trips') ? 'exists' : 'MISSING'; }
  };
  Object.keys(steps).forEach(function(name) {
    try { logEvent_('diag ' + name, steps[name]()); }
    catch (err) { logEvent_('diag ' + name + ' FAILED', err); }
  });
}

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
    logEvent_('doPost', 'action=' + (data.action || 'submit') + ' hasToken=' + !!(data._idToken || data.idToken) + ' chars=' + raw.length);

    // LEADERS ROUTE: verified Google sign-in + email on the Leaders sheet (see Leaders.gs).
    if (data.action === 'addTrip') {
      return jsonOut_(addTrip_(data));
    }

    // PROFILE ROUTES: the user ID comes ONLY from a verified Google ID token.
    // Any googleUserId sent by the client is ignored.
    if (data.action === 'getProfile') {
      var readerId = verifyGoogleToken_(data.idToken);
      if (!readerId) return jsonOut_({ error: 'unauthorized' });
      return jsonOut_(getUserProfile(readerId) || {});
    }

    if (data.action === 'saveProfile') {
      var writerId = verifyGoogleToken_(data.idToken);
      if (!writerId) { logEvent_('saveProfile unauthorized', tokenFailReason_); return jsonOut_({ status: 'error', message: 'unauthorized' }); }
      saveProfileAndYouth_(writerId, sanitizeFields_(data.profileData || {}, []));
      logEvent_('saveProfile ok', 'fields=' + Object.keys(data.profileData || {}).length);
      return jsonOut_({ status: 'success' });
    }

    // SUBMISSION (anonymous, no sign-in required).
    // Soft anti-spam: trim and flag, never reject.
    var flags = [];
    var submittedTripRevision = data._tripRevision;
    var submittedIdToken = data._idToken;
    if (raw.length > MAX_PAYLOAD_CHARS) flags.push('Oversized payload');
    var honeypotFilled = !!(data[HONEYPOT_FIELD] && String(data[HONEYPOT_FIELD]).trim() !== '');
    if (honeypotFilled) flags.push('Honeypot filled');

    var slipSnapshot = data._slipSnapshot; // read before sanitizeFields_ drops '_' keys
    data = sanitizeFields_(data, flags);

    var isAdult = String(data['Form Type']).toLowerCase() === 'adult';
    if (!(isAdult ? data['Adult Name'] : data['Youth Name'])) flags.push('Missing name');
    if (!(isAdult ? data['Adult Signature'] : data['Youth Signature'])) flags.push('Missing signature');

    // Always overwrite server-controlled columns so a client can't spoof them.
    data['tripId'] = displayValue_(trip.tripId); // always from the Trips sheet, never the browser
    data['formVersion'] = FORM_VERSION;
    data['wordingHash'] = wordingHash_(slipSnapshot);
    if (!data['wordingHash']) flags.push('No wording hash');
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

    // 7. Save the signed-in parent's profile (never blocks the submission).
    if (!submittedIdToken) logEvent_('profile skipped', 'no _idToken in submission (not signed in, or old page cached)');
    if (submittedIdToken && !honeypotFilled) {
      try {
        var profileId = verifyGoogleToken_(submittedIdToken);
        if (!profileId) logEvent_('profile skipped: token rejected', tokenFailReason_);
        if (profileId) {
          var profileData = {};
          Object.keys(data).forEach(function(k) {
            if (k === 'Form Type' || k === 'tripId' || k === 'formVersion' || k === 'wordingHash' || k.toLowerCase().indexOf('signature') !== -1) return;
            profileData[k] = data[k];
          });
          saveProfileAndYouth_(profileId, profileData);
          logEvent_('profile saved', 'fields=' + Object.keys(profileData).length);
        }
      } catch (profErr) {
        console.error('Profile save failed: ' + profErr);
        logEvent_('profile save FAILED', profErr + ' | ' + (profErr && profErr.stack || ''));
      }
    }

    return jsonOut_({ status: "success" });

  } catch (err) {
    return jsonOut_({ status: "error", message: err.toString() });
  }
}

// ---------------------------------------------------------------------------
// AUTH + INPUT HYGIENE
// ---------------------------------------------------------------------------

// Client IDs whose tokens are accepted. The test client is only accepted when the
// Script Property ALLOW_TEST_CLIENT is 'true', so production stays locked to one client.
function getAllowedClientIds_() {
  var props = PropertiesService.getScriptProperties();
  var ids = [props.getProperty('OAUTH_CLIENT_ID') || DEFAULT_OAUTH_CLIENT_ID];
//  if (props.getProperty('ALLOW_TEST_CLIENT') === 'true') ids.push(TEST_OAUTH_CLIENT_ID);
  return ids;
}

// Returns the verified Google user ID (the token's 'sub'), or null if the token is invalid.
function verifyGoogleToken_(idToken) {
  tokenFailReason_ = '';
  if (!idToken || typeof idToken !== 'string') { tokenFailReason_ = 'missing token'; return null; }
  try {
    var cache = CacheService.getScriptCache();
    var key = 'tok_' + Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
    var hit = cache.get(key);
    if (hit) return hit;

    var res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) { tokenFailReason_ = 'tokeninfo HTTP ' + res.getResponseCode() + ' ' + res.getContentText().substring(0, 200); return null; }

    var info = JSON.parse(res.getContentText());
    if (getAllowedClientIds_().indexOf(info.aud) === -1) { tokenFailReason_ = 'aud mismatch: token aud=' + info.aud + ' expected one of=' + getAllowedClientIds_().join(','); return null; }     // must be minted for THIS client
    if (Number(info.exp) * 1000 < Date.now()) { tokenFailReason_ = 'token expired'; return null; }  // not expired
    if (!info.sub) return null;

    cache.put(key, String(info.sub), 300);
    return String(info.sub);
  } catch (err) {
    console.error('Token verification failed: ' + err);
    tokenFailReason_ = 'exception: ' + err;
    return null;
  }
}

// Returns a cleaned copy: drops internal/server-controlled keys, caps key count and
// value length (flagging when it trims), and flattens arrays/objects to strings.
function sanitizeFields_(input, flags) {
  var blocked = { 'action': true, 'idToken': true, 'googleUserId': true,
                  'Flags': true, 'PDF Link': true, 'tripId': true, 'formVersion': true, 'wordingHash': true };
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

// Validates the client's snapshot of the page the signer saw, and returns a SHA-256 hash of
// its WORDING only (all fixed text, labels and trip details; not what was typed or which
// radio was ticked). Rows with the same hash were shown identical text. Returns '' if the
// snapshot is missing or malformed.
function wordingHash_(raw) {
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > 400) return '';
  var lines = [];
  for (var i = 0; i < raw.length; i++) {
    var line = raw[i];
    if (!line || !Array.isArray(line.segs) || line.segs.length > 60) return '';
    var text = '';
    for (var j = 0; j < line.segs.length; j++) {
      var sg = line.segs[j];
      if (!sg || typeof sg.s !== 'string') return '';
      if (sg.v) { text += '\u0001'; continue; } // typed value: marks the blank, not its content
      text += sg.s.substring(0, 2000).replace(/^\[[X ]\] $/, '[ ] '); // radios: ignore selection
    }
    lines.push((Number(line.h) || 0) + ':' + text);
  }
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, lines.join('\n'), Utilities.Charset.UTF_8)
    .map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
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
        profile.youthProfiles = getYouthProfiles_(id);
        return profile;
      }
    }
    var youth = getYouthProfiles_(id);
    return youth.length ? { youthProfiles: youth } : null; // Profile not found
  } catch (err) {
    console.error("Error fetching profile: " + err);
    return null;
  }
}

// Save or update profile for a (verified) Google User ID
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
}


// ---------------------------------------------------------------------------
// MULTIPLE YOUTH PER PARENT
// 'Youth ...' fields are stored one row per scout in a 'Youth Profiles' sheet
// (key: Google ID + Youth Name). Everything else stays in the single-row
// 'Profiles' sheet.
// ---------------------------------------------------------------------------

function saveProfileAndYouth_(googleUserId, fields) {
  var parentFields = {};
  var youthFields = {};
  Object.keys(fields).forEach(function(k) {
    if (k.indexOf('Youth ') === 0) youthFields[k] = fields[k];
    else parentFields[k] = fields[k];
  });
  saveOrUpdateUserProfile(googleUserId, parentFields);
  saveYouthProfile_(googleUserId, youthFields);
}

function saveYouthProfile_(googleUserId, youthFields) {
  var name = String(youthFields['Youth Name'] || '').trim();
  if (!name) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Youth Profiles');
    if (!sheet) {
      sheet = ss.insertSheet('Youth Profiles');
      sheet.appendRow(['Google ID']);
      sheet.getRange(1, 1).setFontWeight('bold');
    }
    var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
      .map(function(h) { return h.toString().trim(); });
    Object.keys(youthFields).forEach(function(key) {
      if (headers.indexOf(key) === -1) {
        sheet.getRange(1, headers.length + 1).setValue(key).setFontWeight('bold');
        headers.push(key);
      }
    });

    var id = String(googleUserId);
    var nameCol = headers.indexOf('Youth Name');
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id &&
          String(data[i][nameCol]).trim().toLowerCase() === name.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    var existing = rowIndex > 0 ? data[rowIndex - 1] : [];
    var row = headers.map(function(h, idx) {
      if (h === 'Google ID') return id;
      var v = youthFields[h];
      if (v !== undefined && v !== null) return v;
      return existing[idx] !== undefined ? existing[idx] : '';
    });
    var target = rowIndex > 0 ? rowIndex : sheet.getLastRow() + 1;
    sheet.getRange(target, 1).setNumberFormat('@');
    sheet.getRange(target, 1, 1, row.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
}

function getYouthProfiles_(googleUserId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Youth Profiles');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var id = String(googleUserId);
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== id) continue;
      var y = {};
      for (var j = 1; j < headers.length; j++) {
        if (headers[j] === '') continue;
        var v = data[i][j];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        y[headers[j]] = v;
      }
      out.push(y);
    }
    return out;
  } catch (err) {
    console.error('Error fetching youth profiles: ' + err);
    return [];
  }
}
