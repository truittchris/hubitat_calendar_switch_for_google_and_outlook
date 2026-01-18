# Setup

This guide covers installation and first run for Hubitat Calendar Switch.

Contents

- Installation
- Provider setup
  - Google Calendar OAuth
  - Microsoft (Outlook / Microsoft 365) OAuth
- Create your first switch

## Installation

Option A - Hubitat Package Manager (when available)

- In HPM, install the Hubitat Calendar Switch package.
- Confirm you see both of the following under Drivers Code and Apps Code:
  - App: Hubitat Calendar Switch
  - Driver: Hubitat Calendar Switch - Control Device

Option B - Manual install

1) In your Hubitat hub UI, go to Apps Code and click New App.
2) Paste the contents of:
   - Apps/Hubitat_Calendar_Switch.groovy
3) Click Save.
4) Go to Drivers Code and click New Driver.
5) Paste the contents of:
   - Drivers/Hubitat_Calendar_Switch_Control_Device.groovy
6) Click Save.
7) Go to Apps and click Add User App.
8) Select Hubitat Calendar Switch.

## Provider setup

The app supports two providers. You can configure one or both.

Important

- Keep your client secrets private.
- Use a dedicated OAuth app registration per Hubitat hub, or track which hub is using which credentials.

### Google Calendar OAuth

1) In Google Cloud Console, create or select a project.
2) Enable the Google Calendar API.
3) Configure the OAuth consent screen (External is fine for personal use). Add yourself as a test user if needed.
4) Create OAuth client credentials:
   - Application type: Web application
   - Authorized redirect URI:
     - https://cloud.hubitat.com/oauth/stateredirect
5) Copy the Client ID and Client Secret.
6) In the Hubitat Calendar Switch app, paste the Google client ID and secret.
7) Click Authorize Google and complete the Google consent flow.

Notes

- You may see a Google warning that the app is not verified. This is normal for apps in testing. Continue if you recognize the developer and you created the OAuth client.
- The app stores a refresh token (offline access) so it can keep working without re-authorizing every hour.

### Microsoft (Outlook / Microsoft 365) OAuth

1) In Azure Portal, go to App registrations and create a new registration.
2) Supported account types:
   - Most users should use Accounts in any organizational directory and personal Microsoft accounts.
3) Add a Redirect URI:
   - Platform: Web
   - Redirect URI:
     - https://cloud.hubitat.com/oauth/stateredirect
4) Create a client secret and copy it now.
5) API permissions:
   - Microsoft Graph
   - Delegated permissions:
     - Calendars.Read
     - offline_access
     - User.Read
6) Copy the Application (client) ID.
7) In the Hubitat Calendar Switch app, paste:
   - Microsoft client ID
   - Microsoft client secret
   - Tenant (usually common)
8) Click Authorize Microsoft and complete the Microsoft sign-in and consent flow.

Notes

- If you use a specific tenant, set the Tenant field to that tenant ID.
- If you change permissions, you may need to Disconnect Microsoft and re-authorize.

## Create your first switch

1) In the app, select a provider and enter a switch name.
2) Click Add switch.
3) Open the newly created device.
4) On Preferences, set Match words and Ignore words (optional), plus timing and filters.
5) On Commands, click Test Now.
6) Confirm Current States show the expected active or next event details.
