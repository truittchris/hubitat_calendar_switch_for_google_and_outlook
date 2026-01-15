truittchris Calendar OAuth Bridge v0.1.0

What you are installing
1) App (parent): Calendar OAuth Bridge
2) Driver (child): Calendar Control Switch

This uses Hubitat’s OAuth state redirect endpoint:
https://cloud.hubitat.com/oauth/stateredirect

Installation
1) Hubitat – Apps Code – New App
   - Paste apps/Calendar_OAuth_Bridge.groovy
   - Save
   - Enable the OAuth toggle for the app
2) Hubitat – Drivers Code – New Driver
   - Paste drivers/Calendar_Control_Switch.groovy
   - Save
3) Hubitat – Apps – Add User App
   - Select Calendar OAuth Bridge
   - Configure provider, credentials, and polling

Google Cloud Console setup (Calendar read-only)
1) Create or select a Google Cloud project
2) Enable the Google Calendar API for the project
3) Configure OAuth consent screen (External is fine; add yourself as a test user if needed)
4) Create Credentials – OAuth client ID
   - Application type: Web application
   - Authorized redirect URI:
     https://cloud.hubitat.com/oauth/stateredirect
5) Copy Client ID and Client Secret into the Hubitat app, then Authenticate

Microsoft Entra ID (Azure) setup (Graph Calendar read-only)
1) Create an app registration
2) Add a Redirect URI
   - Platform: Web
   - Redirect URI:
     https://cloud.hubitat.com/oauth/stateredirect
3) API permissions
   - Microsoft Graph (delegated): Calendars.Read
   - offline_access is requested via scope; no separate permission needed in most tenants
4) Create a client secret
5) Copy Tenant (optional), Client ID, and Client Secret into the Hubitat app, then Authenticate
   - Tenant can be: common, organizations, consumers, or a specific tenant ID

Using it
- After authentication, go to Manage child switches
- Add one or more switches, each with:
  - minutes before start to turn on
  - minutes after end to turn off
  - include/exclude keywords (title match, case-insensitive)
  - optional busy/private/all-day filters
- The app polls on the chosen interval and updates the child devices.

Notes and current limitations (v0.1.0)
- Primary calendar only
- Keyword matching is applied to event title only
- No attachments, no event bodies
- Polling is periodic (not push/webhook)

Troubleshooting
- If you see “OAuth is not enabled for this app”, open Apps Code and enable OAuth, then Save.
- If you see provider “redirect_uri_mismatch”, verify the redirect URI is exactly:
  https://cloud.hubitat.com/oauth/stateredirect
- Turn on Debug authentication to log API calls and token lifecycle.

Roadmap ideas (if you want them next)
- Per-switch field matching (subject vs location vs organizer)
- Per-switch calendar selection (not just primary)
- Push updates (Microsoft Graph subscriptions or Google push channels) with a lightweight relay service
