const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwWHDrP4c-V3IQ8lrb8xIF8uUl36InF7PcU_arJ5Lkh1I97W79LlQr6XbIIAy_c7ls5/exec";
const OAUTH_CLIENT_ID = "1059045077738-ohkb0nupq7kh69kj3i5io9n1rbjmfn8m.apps.googleusercontent.com";

let leaderIdToken = null;

function setMessage(text, isError) {
  const el = document.getElementById('leaderMessage');
  el.textContent = text;
  el.style.color = isError ? '#b00020' : '#1b5e20';
}

function parseJwt(token) {
  const base64 = token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
  return JSON.parse(window.atob(base64));
}

function handleLeaderCredential(response) {
  leaderIdToken = response.credential;
  const claims = parseJwt(response.credential);
  document.getElementById('signedInAs').textContent = 'Signed in as ' + claims.email;
  document.getElementById('addTripBtn').disabled = false;
  setMessage('', false);
}

window.addEventListener('load', function() {
  if (typeof google !== 'undefined' && google.accounts && google.accounts.id) {
    google.accounts.id.initialize({ client_id: OAUTH_CLIENT_ID, callback: handleLeaderCredential });
    google.accounts.id.renderButton(document.getElementById('leaderSignInBtn'), { theme: 'outline', size: 'large' });
  }
  document.getElementById('tripForm').addEventListener('submit', submitTrip);
});

function submitTrip(e) {
  e.preventDefault();
  if (!leaderIdToken) { setMessage('Please sign in first.', true); return; }

  const form = e.target;
  const btn = document.getElementById('addTripBtn');
  const payload = { action: 'addTrip', idToken: leaderIdToken };
  new FormData(form).forEach((value, key) => { payload[key] = value; });

  btn.disabled = true;
  setMessage('Adding trip...', false);

  fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) })
    .then(res => res.json())
    .then(r => {
      if (r.status === 'success') {
        setMessage('Trip added. The permission slip now shows this trip.', false);
        form.reset();
      } else if (r.code === 'signin') {
        leaderIdToken = null;
        document.getElementById('signedInAs').textContent = '';
        setMessage(r.message + ' Your entries are still in the form.', true);
        return;
      } else {
        setMessage(r.message || 'Could not add the trip.', true);
      }
    })
    .catch(err => {
      console.error(err);
      setMessage('Network error. Your entries are still in the form; try again.', true);
    })
    .finally(() => { btn.disabled = !leaderIdToken; });
}
