const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwWHDrP4c-V3IQ8lrb8xIF8uUl36InF7PcU_arJ5Lkh1I97W79LlQr6XbIIAy_c7ls5/exec";
const DRAFT_KEY = 'troop58_form_draft';

document.addEventListener('DOMContentLoaded', function() {
  loadTripDetails();
  restoreFormDraft();

  // Force the toggle to run on page load to strip 'required' from hidden fields
  toggleFormType();

  const form = document.getElementById('permissionForm');
  if (form) {
    form.addEventListener('input', saveFormDraft);
    form.addEventListener('change', saveFormDraft);
    form.addEventListener('submit', handleFormSubmit);
  }
});

// Fetch dynamic trip configuration from Google Sheet API
function loadTripDetails() {
  fetch(APPS_SCRIPT_URL)
    .then(res => res.json())
    .then(data => {
      // Loop over every key returned from the Spreadsheet / Apps Script API
      Object.keys(data).forEach(key => {
        const val = data[key];
        if (val === undefined || val === null) return;

        const element = document.getElementById(key);
        if (element) {
          // If the payload value contains HTML (e.g. custom departure radio buttons), inject innerHTML
          if (typeof val === 'string' && (val.includes('<') && val.includes('>'))) {
            element.innerHTML = val;
          } else if ('value' in element && (element.tagName === 'INPUT' || element.tagName === 'TEXTAREA' || element.tagName === 'SELECT')) {
            element.value = val;
          } else {
            element.innerText = val;
          }
        }
      });
    })
    .catch(err => console.error("Error loading trip data:", err));
}

// Submit form data to Google Sheet API
function handleFormSubmit(e) {
  e.preventDefault();
  const form = e.target;
  const formData = new FormData(form);
  const payload = {};

  formData.forEach((value, key) => {
    if (payload[key]) {
      if (!Array.isArray(payload[key])) payload[key] = [payload[key]];
      payload[key].push(value);
    } else {
      payload[key] = value;
    }
  });

  const submitBtn = form.querySelector('button[type="submit"]');
  if (submitBtn) submitBtn.disabled = true;

  fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    body: JSON.stringify(payload)
  })
  .then(res => res.json())
  .then(response => {
    if (response.status === 'success') {
      alert("Permission slip submitted successfully! Thank you.");
      if (window.currentUserGoogleId) saveUserProfile(window.currentGoogleUserId); // Save profile for future auto-fill
      localStorage.removeItem(DRAFT_KEY);
      form.reset();
      toggleFormType();
    } else {
      alert("Submission error: " + response.message);
    }
  })
  .catch(err => {
    alert("Network error submitting form. Please try again.");
    console.error(err);
  })
  .finally(() => {
    if (submitBtn) submitBtn.disabled = false;
  });
}

// Youth/Adult Form Switcher
function toggleFormType() {
  var formType = document.getElementById('formType').value.toLowerCase();
  var youthSec = document.getElementById('youthSection');
  var adultSec = document.getElementById('adultSection');

  if (formType === 'youth') {
    youthSec.style.display = 'block';
    adultSec.style.display = 'none';
    toggleRequired(youthSec, true);
    toggleRequired(adultSec, false);
  } else {
    youthSec.style.display = 'none';
    adultSec.style.display = 'block';
    toggleRequired(youthSec, false);
    toggleRequired(adultSec, true);
  }
}

function toggleRequired(container, enable) {
  var inputs = container.querySelectorAll('input, select, textarea');
  inputs.forEach(function(input) {
    if (enable) {
      if (input.dataset.wasRequired === 'true') input.required = true;
    } else {
      if (input.required) {
        input.dataset.wasRequired = 'true';
        input.required = false;
      }
    }
  });
}

function saveFormDraft() {
  const form = document.getElementById('permissionForm');
  if (!form) return;
  const data = extractFormData(form, { skipSignatures: false });
  data._tripId = window.currentTripId || null;
  localStorage.setItem(DRAFT_KEY, JSON.stringify(data));
}

function restoreFormDraft() {
  const savedData = localStorage.getItem(DRAFT_KEY);
  if (!savedData) return;
  try {
    const data = JSON.parse(savedData);
    if (data._tripId && window.currentTripId && data._tripId !== window.currentTripId) {
      localStorage.removeItem(DRAFT_KEY);
      return;
    }
    const form = document.getElementById('permissionForm');
    populateFormFields(form, data, { skipSignatures: false });
  } catch (e) {
    console.error(e);
  }
}

// Initialize GIS on page load
window.onload = function () {
  google.accounts.id.initialize({
    client_id: "1059045077738-ohkb0nupq7kh69kj3i5io9n1rbjmfn8m.apps.googleusercontent.com",
    callback: handleCredentialResponse
  });

  google.accounts.id.renderButton(
    document.getElementById("googleSignInBtn"),
    { theme: "outline", size: "large" }
  );
};

// Handle successful sign-in
function handleCredentialResponse(response) {
  // Decode JWT token payload
  const responsePayload = parseJwt(response.credential);
  
  // 1. Auto-fill contact details
  document.getElementById("parentName").value = responsePayload.name;
  document.getElementById("parentEmail").value = responsePayload.email;

  // 2. Retrieve & load saved user profile from localStorage or backend
  loadUserProfile(responsePayload.sub);

  // 3. ENFORCE CLEAN SIGNATURE STATE
  // Explicitly clear signature fields so typing a name is required for every submission
  const sigInput = document.getElementById("parentSignature");
  if (sigInput) sigInput.value = "";
}

function saveUserProfile(googleUserId) {
  if (!googleUserId) return;
  const form = document.getElementById('permissionForm');
  const profileData = extractFormData(form, { skipSignatures: true });

  fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    body: JSON.stringify({
      action: 'saveProfile',
      googleUserId: googleUserId,
      profileData: profileData
    })
  }).catch(err => console.error("Profile save error:", err));
}

// Simple JWT parser helper
function parseJwt(token) {
  const base64Url = token.split('.')[1];
  const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
  return JSON.parse(window.atob(base64));
}

// Updated loadUserProfile to fetch from Backend API for multi-device support
function loadUserProfile(googleUserId) {
  if (!googleUserId) return;

  // 1. Try local cache first for instant loading
  const localCache = localStorage.getItem(`user_profile_${googleUserId}`);
  if (localCache) {
    populateFormFields(document.getElementById('permissionForm'), JSON.parse(localCache), { skipSignatures: true });
    return; // Exit early if local cache is found
  }

  // 2. Fetch fresh profile from Google Apps Script backend (Multi-Device Sync)
  fetch(`${APPS_SCRIPT_URL}?action=getProfile&googleUserId=${googleUserId}`)
    .then(res => res.json())
    .then(profileData => {
      if (profileData && !profileData.error) {
        // Populate form with cross-device data
        populateFormFields(document.getElementById('permissionForm'), profileData, { skipSignatures: true });
        
        // Update local cache
        localStorage.setItem(`user_profile_${googleUserId}`, JSON.stringify(profileData));
        
        // Keep active draft updated
        saveFormDraft(window.currentTripId);
      }
    })
    .catch(err => console.log('Offline or profile fetch error:', err));
}

// Render the Scout selection UI dynamically
function renderYouthSelector(youthProfiles) {
  let selector = document.getElementById('scoutSelector');
  const youthContainer = document.getElementById('scoutSelectorContainer');

  if (!youthProfiles || youthProfiles.length === 0) {
    if (youthContainer) youthContainer.style.display = 'none';
    return;
  }

  // If container doesn't exist, create it above the Scout Name field
  if (!selector) {
    const html = `
      <div id="scoutSelectorContainer" class="form-group" style="background: #eef5fc; padding: 12px; border-radius: 6px; margin-bottom: 15px;">
        <label for="scoutSelector"><strong>Select Scout to Pre-fill:</strong></label>
        <select id="scoutSelector" onchange="onScoutSelected(this.value)" style="width: 100%; padding: 8px; margin-top: 5px;">
          <option value="">-- Choose a Scout --</option>
          ${youthProfiles.map((y, idx) => `<option value="${idx}">${y.scoutName} (${y.patrol || 'Troop 58'})</option>`).join('')}
          <option value="new">+ Add Another Scout</option>
        </select>
      </div>
    `;
    const scoutSection = document.getElementById('youthSection') || document.getElementById('scoutName').parentNode;
    scoutSection.insertAdjacentHTML('afterbegin', html);
    selector = document.getElementById('scoutSelector');
  }

  // Cache profiles in memory for quick toggling
  window.userYouthProfiles = youthProfiles;
}

// Handle switching between Scouts
function onScoutSelected(selectedIndex) {
  if (selectedIndex === "" || selectedIndex === "new") {
    // Clear Scout fields for a fresh entry
    document.getElementById('scoutName').value = "";
    if (document.getElementById('patrol')) document.getElementById('patrol').value = "";
    if (document.getElementById('dietary')) document.getElementById('dietary').value = "";
    return;
  }

  const selectedScout = window.userYouthProfiles[selectedIndex];
  if (!selectedScout) return;

  // Auto-fill Scout fields
  if (document.getElementById('scoutName')) document.getElementById('scoutName').value = selectedScout.scoutName || "";
  if (document.getElementById('patrol')) document.getElementById('patrol').value = selectedScout.patrol || "";
  if (document.getElementById('dietary')) document.getElementById('dietary').value = selectedScout.dietary || "";
  if (document.getElementById('medical')) document.getElementById('medical').value = selectedScout.medical || "";

  // Trigger local draft update so refreshing remembers the currently selected Scout
  saveFormDraft(window.currentTripId);
}

// Helper to identify signature/date fields dynamically
function isSignatureOrDateField(key) {
  const k = key.toLowerCase();
  return k.includes('signature') || k.includes('sigdate') || k.includes('sign_date');
}

// Dynamically extracts form data into an object
function extractFormData(form, options = { skipSignatures: false }) {
  const formData = new FormData(form);
  const data = {};
  formData.forEach((value, key) => {
    if (options.skipSignatures && isSignatureOrDateField(key)) return;
    if (data[key]) {
      if (!Array.isArray(data[key])) data[key] = [data[key]];
      data[key].push(value);
    } else {
      data[key] = value;
    }
  });
  return data;
}

// Dynamically populates form inputs from an object
function populateFormFields(form, data, options = { skipSignatures: false }) {
  Object.keys(data).forEach(key => {
    if (key.startsWith('_')) return; 
    if (options.skipSignatures && isSignatureOrDateField(key)) return;

    const elements = form.querySelectorAll(`[name="${key}"]`);
    const value = data[key];

    elements.forEach(el => {
      if (el.type === 'checkbox' || el.type === 'radio') {
        el.checked = Array.isArray(value) ? value.includes(el.value) : (el.value === value);
      } else {
        el.value = value;
      }
    });
  });

  // TRIGGER YOUTH SELECTOR IF ARRAY EXISTS
  if (profileData.youthProfiles) {
    renderYouthSelector(profileData.youthProfiles);
  }
}