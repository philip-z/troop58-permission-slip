# troop58-permission-slip

A lightweight web form for collecting Troop 58 trip permission slips. The app is built as a static HTML/CSS/JavaScript form that can be served from a simple web host, while submitting data to a Google Apps Script endpoint and saving drafts locally in the browser.

## Overview

This project creates a digital permission slip for youth and adult participants. It supports:

- Youth and adult form variants
- Dynamic trip details loaded from a Google Sheet / Apps Script endpoint
- Google Sign-In for user identification and profile auto-fill
- Local draft persistence so users can recover partially completed submissions
- Submission of completed forms to a Google Apps Script backend
- Optional profile caching for repeated form use

## Features

### Form mode switching
The form includes a selector for either:

- Youth Permission Slip
- Adult Permission Slip

The JavaScript toggles required fields based on the selected form type so the correct inputs are enforced without leaving hidden fields active.

### Dynamic trip configuration
Trip fields such as:

- trip title
- location
- dates
- cost
- departure and return details
- coordinator
- uniform

are loaded at runtime from the configured Apps Script endpoint. This keeps the form content easy to update without changing the HTML.

### Google sign-in and profile support
The page initializes Google Identity Services and reads the signed-in user profile to:

- prefill contact information when available
- save a reusable profile for future submissions
- restore stored profile information across devices via the backend

### Draft saving
The app automatically saves the current form state to `localStorage` and restores it on page load if the same trip is still active. This reduces lost work during refreshes or interruptions.

## Project structure

- `index.html` — main web form markup
- `style.css` — page styling and layout
- `script.js` — form logic, validation toggling, sign-in flow, draft saving, and form submission
- `images/` — troop branding assets

## Local development

Because this is a static app, you can run it in any basic web server environment.

### Option 1: Open directly
Open `index.html` in a browser.

### Option 2: Use a local server
From the project root:

```bash
python -m http.server 8000
```

Then open:

```text
http://localhost:8000
```

## Required configuration

The app depends on a Google Apps Script endpoint configured in `script.js`:

```javascript
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/....../exec";
```

This URL is the backend that:

- returns trip details
- accepts submitted permission slips
- optionally stores user profiles

You will need to replace the sample URL with your live Apps Script deployment.

The Google OAuth client ID is also configured in `script.js`:

```javascript
google.accounts.id.initialize({
  client_id: "YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com",
  callback: handleCredentialResponse
});
```

Update this to match the OAuth client used by your Google project.

## Submission flow

When the user submits the form:

1. Form data is collected from the HTML form.
2. The data is serialized into a payload.
3. A `POST` request is sent to the Apps Script endpoint.
4. The backend responds with a success or error status.
5. The page displays an alert and clears the local draft on success.

## Draft behavior

Drafts are stored in `localStorage` under the key:

```javascript
const DRAFT_KEY = 'troop58_form_draft';
```

This allows users to reopen the form and continue where they left off, as long as the trip remains the same.

## Notes

- This project is designed for a specific troop workflow and is tightly coupled to the Google Apps Script backend.
- The HTML structure and JavaScript are configured for the troop's permission slip fields and wording.
- If the Apps Script backend changes field names or payload structure, both the frontend and backend must stay in sync.

## Deployment

This application can be deployed as a static website on any basic hosting provider. The main requirement is that the Google Apps Script URL and OAuth client ID remain configured for the production environment.

## Future enhancements

Potential improvements include:

- PDF generation for submitted forms
- stronger validation rules for required fields
- admin dashboard for reviewing submissions
- improved data export and reporting

