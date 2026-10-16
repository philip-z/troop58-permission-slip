const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwWHDrP4c-V3IQ8lrb8xIF8uUl36InF7PcU_arJ5Lkh1I97W79LlQr6XbIIAy_c7ls5/exec";
const DRAFT_KEY = 'troop58_form_draft';

document.addEventListener('DOMContentLoaded', function() {
  loadTripDetails();
  restoreFormDraft();

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
      if (data.tripTitle) document.getElementById('tripTitle').innerText = data.tripTitle;
      if (data.location) document.getElementById('tripLocation').innerText = data.location;
      if (data.dates) document.getElementById('tripDates').innerText = data.dates;
      if (data.returnDetails) document.getElementById('returnDetails').innerText = data.returnDetails;
      if (data.cost) document.getElementById('tripCost').innerText = data.cost;
      if (data.coordinator) document.getElementById('tripCoordinator').innerText = data.coordinator;
      if (data.uniform) document.getElementById('tripUniform').innerText = data.uniform;
      
      if (data.departureHtml) {
        document.getElementById('departureContainer').innerHTML = data.departureHtml;
      }
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
  var formType = document.getElementById('formType').value;
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
  const formData = new FormData(form);
  const data = {};
  formData.forEach((value, key) => { data[key] = value; });
  localStorage.setItem(DRAFT_KEY, JSON.stringify(data));
}

function restoreFormDraft() {
  const savedData = localStorage.getItem(DRAFT_KEY);
  if (!savedData) return;
  try {
    const data = JSON.parse(savedData);
    const form = document.getElementById('permissionForm');
    Object.keys(data).forEach(key => {
      const el = form.querySelector(`[name="${key}"]`);
      if (el) el.value = data[key];
    });
    toggleFormType();
  } catch (e) {
    console.error(e);
  }
}