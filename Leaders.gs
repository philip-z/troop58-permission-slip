// ---------------------------------------------------------------------------
// LEADERS: add a new trip to the Trips sheet.
// Access = a valid Google sign-in whose verified email is listed in column A
// of a sheet tab named "Leaders". Everything is enforced here on the server;
// the leaders.html page itself is public.
// ---------------------------------------------------------------------------

// Must match the Trips sheet headers. doGet publishes every Trips column
// publicly, so NEVER add private columns (like who added a trip) to that sheet.
var TRIP_COLUMNS = ['tripTitle', 'tripLocation', 'tripDates', 'departureDetails',
                    'returnDetails', 'tripCost', 'tripCoordinator', 'tripUniform',
                    'customDepartureHtml'];
var TRIP_REQUIRED = ['tripTitle', 'tripLocation', 'tripDates'];
var TRIP_FIELD_MAX = 500;
var MAX_DEPARTURE_OPTIONS = 20;
var DEPARTURE_OPTION_MAX = 100;

// Returns { sub, email, emailVerified } for a valid token minted for our client, else null.
function verifyGoogleTokenInfo_(idToken) {
  if (!idToken || typeof idToken !== 'string') return null;
  try {
    var cache = CacheService.getScriptCache();
    var key = 'tokinfo_' + Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
    var hit = cache.get(key);
    if (hit) return JSON.parse(hit);

    var res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;

    var info = JSON.parse(res.getContentText());
    if (info.aud !== getOAuthClientId_()) return null;
    if (Number(info.exp) * 1000 < Date.now()) return null;
    if (!info.sub) return null;

    var out = {
      sub: String(info.sub),
      email: String(info.email || '').trim().toLowerCase(),
      emailVerified: String(info.email_verified) === 'true'
    };
    cache.put(key, JSON.stringify(out), 300);
    return out;
  } catch (err) {
    console.error('Token info verification failed: ' + err);
    return null;
  }
}

function isLeaderEmail_(email) {
  if (!email) return false;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leaders');
  if (!sheet || sheet.getLastRow() < 1) return false;
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === email) return true;
  }
  return false;
}

function cleanTripText_(v) {
  return String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim().substring(0, TRIP_FIELD_MAX);
}

// Builds the customDepartureHtml cell from plain text. Every value is escaped, so
// leaders never paste raw HTML. The public page replaces the departure span's
// contents with this, so it also carries the departure details text.
function buildDepartureHtml_(departureDetails, options) {
  if (!options.length) return '';
  var opts = options.map(function(o) {
    return '<option value="' + escapeHtml_(o) + '">' + escapeHtml_(o) + '</option>';
  }).join('');
  return escapeHtml_(departureDetails) +
    '<br><label for="departureOption"><strong>Departure option:</strong></label> ' +
    '<select id="departureOption" name="Departure Option" required>' +
    '<option value="">-- Choose one --</option>' + opts + '</select><br>';
}

function addTrip_(data) {
  var info = verifyGoogleTokenInfo_(data.idToken);
  if (!info) return { status: 'error', code: 'signin', message: 'Your sign-in has expired or is invalid. Please sign in again.' };
  if (!info.emailVerified || !isLeaderEmail_(info.email)) {
    return { status: 'error', code: 'not_leader', message: 'Your Google account (' + info.email + ') is not on the leaders list.' };
  }

  var trip = {};
  TRIP_COLUMNS.forEach(function(c) { trip[c] = cleanTripText_(data[c]); });
  var missing = TRIP_REQUIRED.filter(function(c) { return !trip[c]; });
  if (missing.length) return { status: 'error', code: 'invalid', message: 'Missing required field(s): ' + missing.join(', ') };

  var options = String(data.departureOptions || '').split(/\r?\n/)
    .map(function(s) { return s.replace(/\s+/g, ' ').trim().substring(0, DEPARTURE_OPTION_MAX); })
    .filter(function(s) { return s !== ''; })
    .slice(0, MAX_DEPARTURE_OPTIONS);
  // Raw HTML from the client is ignored; it is always generated here.
  trip.customDepartureHtml = buildDepartureHtml_(trip.departureDetails, options);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Trips');
    if (!sheet) return { status: 'error', code: 'config', message: 'Trips sheet is missing.' };

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return h.toString().trim(); });
    TRIP_COLUMNS.forEach(function(c) {
      if (headers.indexOf(c) === -1) {
        sheet.getRange(1, headers.length + 1).setValue(c).setFontWeight('bold');
        headers.push(c);
      }
    });

    var row = headers.map(function(h) { return trip[h] !== undefined ? trip[h] : ''; });
    var targetRow = sheet.getLastRow() + 1;
    // Plain text, so Sheets doesn't turn "Nov 14-16" or "$25" into dates/numbers.
    sheet.getRange(targetRow, 1, 1, row.length).setNumberFormat('@').setValues([row]);

    // Who added it goes in a private log tab, never in Trips (doGet publishes Trips).
    var log = ss.getSheetByName('Trip Log') || ss.insertSheet('Trip Log');
    if (log.getLastRow() === 0) log.appendRow(['Timestamp', 'Added By', 'Trip Title', 'Trip Dates']);
    log.appendRow([new Date(), info.email, trip.tripTitle, trip.tripDates]);
  } finally {
    lock.releaseLock();
  }
  return { status: 'success' };
}
