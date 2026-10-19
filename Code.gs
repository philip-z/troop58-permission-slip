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
      saveOrUpdateUserProfile(data.googleUserId, data.profileData);
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
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

// Add to Code.gs: Fetch saved profile by Google User ID dynamically
function getUserProfile(googleUserId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Profiles');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null; // No users saved yet

    // 1. Read exact headers from Row 1
    var headers = data[0].map(function(h) { return h.toString().trim(); });
    
    // 2. Search row by row for matching Google ID in Column A (index 0)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === googleUserId) {
        
        var profile = {};
        // 3. Dynamically map every sheet column to the profile object
        for (var j = 1; j < headers.length; j++) {
          var key = headers[j];
          if (key !== "") {
            profile[key] = data[i][j];
          }
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

// Add to Code.gs: Save or update profile on submission dynamically
function saveOrUpdateUserProfile(googleUserId, profileData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Profiles');
  var headers = [];
  
  // 1. Ensure the Profiles sheet and the ID column exist
  if (!sheet) {
    sheet = ss.insertSheet('Profiles');
    headers = ["Google ID"];
    sheet.appendRow(headers).getRange(1, 1).setFontWeight("bold");
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

  // 3. Find if this user already has a row
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === googleUserId) {
      rowIndex = i + 1; // Apps Script ranges are 1-based, array is 0-based
      break;
    }
  }

  // 4. Build the row payload perfectly matching the dynamic header order
  var rowPayload = headers.map(function(header) {
    if (header === "Google ID") return googleUserId;
    
    return (profileData[header] !== undefined && profileData[header] !== null) 
      ? profileData[header] 
      : "";
  });

  // 5. Update the existing row, or append a new one
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowPayload.length).setValues([rowPayload]);
  } else {
    sheet.appendRow(rowPayload);
  }
}