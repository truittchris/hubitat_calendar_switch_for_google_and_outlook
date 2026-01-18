# Troubleshooting

## OAuth window goes blank after consent

Symptoms

- After Google or Microsoft consent, the browser window appears blank or does not close.

What it means

- The OAuth callback succeeded, but the browser did not render a completion page. The app may already have stored the token.

What to do

- Check the app logs for a token storage message (for example, "Stored Google token").
- Return to the app main page and confirm the provider shows Connected.

## Bad Message 414 (URI Too Long) or Bad Message 431 (Request Header Fields Too Large)

Symptoms

- After returning from OAuth, you see a Jetty error page with 414 or 431.

Common causes

- The browser is reloading a very long Hubitat URL (often via the Back button), including form state and referrer parameters.

What to do

- Do not use the browser Back button from the OAuth flow.
- Close the OAuth tab/window and return to the Hubitat UI.
- Re-open the app from Apps (not from browser history).
- If it persists, clear site data for your hub IP in the browser (cookies and cached data).

## "Microsoft authorization failed - Microsoft did not return an access token"

Most common causes

- Missing delegated permissions (Calendars.Read and offline_access).
- Wrong redirect URI (must be https://cloud.hubitat.com/oauth/stateredirect).
- Tenant mismatch (for example, using a tenant-specific endpoint but Tenant set to common).

What to do

- Verify the redirect URI in Azure exactly matches.
- Verify Graph delegated permissions include Calendars.Read, offline_access, and User.Read.
- Disconnect Microsoft in the app, then authorize again.

## "Device type ... not found" when adding a switch

Meaning

- The driver was not installed, or the driver name and namespace do not match what the app expects.

What to do

- Confirm this driver exists under Drivers Code:
  - Name: Hubitat Calendar Switch - Control Device
  - Namespace: truittchris
- Confirm you clicked Save after creating the driver.
- Re-run Add switch.

## Switch never turns on

Checklist

- Run Test Now on the device.
- Confirm Current States show upcomingSummary or nextEventTitle.
- Confirm your event is within the app look-ahead window.
- If Match words is set, confirm the event title or description contains at least one match word.
- If Ignore words is set, confirm the event does not contain any ignore word.
- If Only consider busy events is enabled, confirm the event is marked Busy in the calendar.

## Getting help

If you open an issue, include:

- Hubitat model and platform version
- App version and driver version
- Provider (Google or Microsoft)
- Relevant app and device logs around the time of the issue