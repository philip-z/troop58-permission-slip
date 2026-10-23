const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwWHDrP4c-V3IQ8lrb8xIF8uUl36InF7PcU_arJ5Lkh1I97W79LlQr6XbIIAy_c7ls5/exec";
const DRAFT_KEY = 'troop58_form_draft';

function updatePrintHeader() {
  const formType = document.getElementById('formType');
  const header = document.getElementById('printHeader');

  if (!formType || !header) return;
  header.textContent = formType.value === 'Adult' ? 'ADULT PERMISSION SLIP' : 'YOUTH PERMISSION SLIP';
}

document.addEventListener('DOMContentLoaded', function() {
  loadTripDetails();
  restoreFormDraft();

  const form = document.getElementById('permissionForm');
  const formType = document.getElementById('formType');

  if (formType) {
    formType.addEventListener('change', function() {
      updatePrintHeader();
      toggleFormType();
    });
  }

  if (form) {
    form.addEventListener('input', saveFormDraft);
    form.addEventListener('change', saveFormDraft);
    form.addEventListener('submit', handleFormSubmit);
  }

  toggleFormType();
  updatePrintHeader();
});

window.addEventListener('load', function() {
  if (typeof google !== 'undefined' && google.accounts && google.accounts.id) {
    google.accounts.id.initialize({
      client_id: "1059045077738-ohkb0nupq7kh69kj3i5io9n1rbjmfn8m.apps.googleusercontent.com",
      callback: handleCredentialResponse
    });

    const signInBtn = document.getElementById('googleSignInBtn');
    if (signInBtn) {
      google.accounts.id.renderButton(signInBtn, { theme: 'outline', size: 'large' });
    }
  }
});

function loadTripDetails() {
  fetch(APPS_SCRIPT_URL)
    .then(res => res.json())
    .then(data => {
      Object.keys(data).forEach(key => {
        const val = data[key];
        if (val === undefined || val === null) return;

        const element = document.getElementById(key);
        if (element) {
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
      if (window.currentUserGoogleId) saveUserProfile(window.currentUserGoogleId);
      localStorage.removeItem(DRAFT_KEY);
      form.reset();
      toggleFormType();
      updatePrintHeader();
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

function toggleFormType() {
  const isYouth = document.getElementById('formType').value === 'Youth';
  setSectionActive(document.getElementById('youthSection'), isYouth);
  setSectionActive(document.getElementById('adultSection'), !isYouth);
}

function setSectionActive(section, active) {
  if (!section) return;
  section.disabled = !active;
  section.hidden = !active;
}

function saveFormDraft() {
  const form = document.getElementById('permissionForm');
  if (!form) return;

  let previous = {};
  try { previous = JSON.parse(localStorage.getItem(DRAFT_KEY)) || {}; } catch (e) {}

  const data = Object.assign(previous, extractFormData(form, { skipSignatures: false }));
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

function handleCredentialResponse(response) {
  const responsePayload = parseJwt(response.credential);
  window.currentUserGoogleId = responsePayload.sub;

  const pName = document.getElementById('parentName');
  const pEmail = document.getElementById('parentEmail');

  if (pName) pName.value = responsePayload.name;
  if (pEmail) pEmail.value = responsePayload.email;

  loadUserProfile(responsePayload.sub);

  document.querySelectorAll('#permissionForm input[name*="signature" i]').forEach(el => {
    el.value = '';
  });
}

function saveUserProfile(googleUserId) {
  if (!googleUserId) return;
  const form = document.getElementById('permissionForm');
  const profileData = extractFormData(form, { skipSignatures: true });
  delete profileData['Form Type'];

  const cacheKey = `user_profile_${googleUserId}`;
  let cached = {};
  try { cached = JSON.parse(localStorage.getItem(cacheKey)) || {}; } catch (e) {}
  localStorage.setItem(cacheKey, JSON.stringify(Object.assign(cached, profileData)));

  fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    body: JSON.stringify({
      action: 'saveProfile',
      googleUserId: googleUserId,
      profileData: profileData
    })
  }).catch(err => console.error("Profile save error:", err));
}

function parseJwt(token) {
  const base64Url = token.split('.')[1];
  const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
  return JSON.parse(window.atob(base64));
}

function loadUserProfile(googleUserId) {
  if (!googleUserId) return;

  const localCache = localStorage.getItem(`user_profile_${googleUserId}`);
  if (localCache && localCache !== '{}') {
    populateFormFields(document.getElementById('permissionForm'), JSON.parse(localCache), { skipSignatures: true });
    return;
  }

  fetch(`${APPS_SCRIPT_URL}?action=getProfile&googleUserId=${googleUserId}`)
    .then(res => res.json())
    .then(profileData => {
      if (profileData && !profileData.error && Object.keys(profileData).length > 0) {
        populateFormFields(document.getElementById('permissionForm'), profileData, { skipSignatures: true });
        localStorage.setItem(`user_profile_${googleUserId}`, JSON.stringify(profileData));
        saveFormDraft(window.currentTripId);
      }
    })
    .catch(err => console.log('Offline or profile fetch error:', err));
}

function renderYouthSelector(youthProfiles) {
  let selector = document.getElementById('scoutSelector');
  const youthContainer = document.getElementById('scoutSelectorContainer');

  if (!youthProfiles || youthProfiles.length === 0) {
    if (youthContainer) youthContainer.style.display = 'none';
    return;
  }

  if (!selector) {
    const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

    const html = `
      <div id="scoutSelectorContainer" class="scout-selector">
        <label for="scoutSelector"><strong>Select Scout to Pre-fill:</strong></label>
        <select id="scoutSelector">
          <option value="">-- Choose a Scout --</option>
          ${youthProfiles.map((y, idx) => `<option value="${idx}">${esc(y.scoutName)} (${esc(y.patrol || 'Troop 58')})</option>`).join('')}
          <option value="new">+ Add Another Scout</option>
        </select>
      </div>
    `;

    const scoutSection = document.getElementById('youthSection');
    if (scoutSection) scoutSection.insertAdjacentHTML('afterend', html);
    selector = document.getElementById('scoutSelector');
    selector.addEventListener('change', e => onScoutSelected(e.target.value));
  }

  window.userYouthProfiles = youthProfiles;
}

function onScoutSelected(selectedIndex) {
  if (selectedIndex === '' || selectedIndex === 'new') {
    const scoutName = document.getElementById('scoutName');
    if (scoutName) scoutName.value = '';
    if (document.getElementById('patrol')) document.getElementById('patrol').value = '';
    if (document.getElementById('dietary')) document.getElementById('dietary').value = '';
    return;
  }

  const selectedScout = window.userYouthProfiles[selectedIndex];
  if (!selectedScout) return;

  if (document.getElementById('scoutName')) document.getElementById('scoutName').value = selectedScout.scoutName || '';
  if (document.getElementById('patrol')) document.getElementById('patrol').value = selectedScout.patrol || '';
  if (document.getElementById('dietary')) document.getElementById('dietary').value = selectedScout.dietary || '';
  if (document.getElementById('medical')) document.getElementById('medical').value = selectedScout.medical || '';

  saveFormDraft(window.currentTripId);
}

function isSignatureOrDateField(key) {
  const k = key.toLowerCase();
  return k.includes('signature') || k.includes('sigdate') || k.includes('sign_date');
}

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

  if (data.youthProfiles) {
    renderYouthSelector(data.youthProfiles);
  }
}
