Hubitat Calendar Switch (OAuth) – calendar-driven control switches for Google and Microsoft



Hubitat Calendar Switch is a Hubitat app plus a child switch driver. It connects to Google Calendar and/or Microsoft 365 (Outlook) using OAuth, fetches events on a schedule, and delivers a normalized event list to each child switch. Each switch contains its own match rules (must include words, ignore words, timing, and filters) and sets its own On/Off state.



What you are installing

1\) App (parent): Hubitat Calendar Switch

2\) Driver (child): Hubitat Calendar Switch – Control Device



How it works

\- You authenticate the app to Google and/or Microsoft.

\- The app fetches calendar events on a schedule (example: every 5 minutes).

\- Each child switch re-evaluates every minute using the most recently fetched events.

\- Each child switch turns On when an event matches its rules and turns Off when no matching event is active (including your before/after buffers).



Prerequisites

\- Hubitat Elevation hub

\- A Google account (for Google Calendar) and/or Microsoft account (for Microsoft 365 / Outlook)

\- Ability to create OAuth credentials:

&nbsp; - Google Cloud Console project access, and/or

&nbsp; - Microsoft Entra ID (Azure) app registration access

\- This project uses Hubitat’s OAuth state redirect endpoint:

&nbsp; https://cloud.hubitat.com/oauth/stateredirect



Install (Hubitat)

Step 1 – Install the driver

1\) Hubitat – Drivers Code – New Driver

2\) Paste the contents of:

&nbsp;  drivers/Hubitat\_Calendar\_Switch\_-\_Control\_Device.groovy

3\) Save



Step 2 – Install the app

1\) Hubitat – Apps Code – New App

2\) Paste the contents of:

&nbsp;  apps/Hubitat\_Calendar\_Switch.groovy

3\) Save

4\) Important: Enable the OAuth toggle for the app (in Apps Code), then Save again



Step 3 – Add the app

1\) Hubitat – Apps – Add User App

2\) Select: Hubitat Calendar Switch

3\) Configure and authenticate (next sections)



Google setup (get Client ID and Client Secret)

Goal: create a Google OAuth Client ID + Secret that can read calendars (read-only).



1\) Create or select a Google Cloud project

\- Go to Google Cloud Console and select an existing project or create a new one.



2\) Enable the Google Calendar API

\- APIs \& Services – Library – search for “Google Calendar API” – Enable



3\) Configure the OAuth consent screen

\- APIs \& Services – OAuth consent screen

\- User type: External is fine for most personal use

\- Add yourself as a Test user if Google requires it

\- Save



4\) Create OAuth credentials

\- APIs \& Services – Credentials – Create Credentials – OAuth client ID

\- Application type: Web application

\- Authorized redirect URI (must match exactly):

&nbsp; https://cloud.hubitat.com/oauth/stateredirect

\- Create



5\) Copy credentials into Hubitat

\- Copy the Client ID and Client Secret

\- In the Hubitat Calendar Switch app:

&nbsp; - Paste Google client ID

&nbsp; - Paste Google client secret

&nbsp; - Click Authorize Google



If you see redirect\_uri\_mismatch

\- Verify the redirect URI is exactly:

&nbsp; https://cloud.hubitat.com/oauth/stateredirect

\- No trailing slash, no http, no extra parameters



Microsoft setup (get Client ID and Client Secret)

Goal: register an app in Microsoft Entra ID and allow delegated read-only calendar access.



1\) Create an app registration

\- Go to Microsoft Entra admin center (Azure)

\- Entra ID – App registrations – New registration

\- Name: Hubitat Calendar Switch (or similar)

\- Supported account types:

&nbsp; - For most people, “Accounts in any organizational directory and personal Microsoft accounts” works well

\- Register



2\) Add a redirect URI

\- In the app registration: Authentication – Add a platform – Web

\- Redirect URI (must match exactly):

&nbsp; https://cloud.hubitat.com/oauth/stateredirect

\- Save



3\) Add Microsoft Graph permissions

\- API permissions – Add a permission – Microsoft Graph – Delegated permissions

\- Add:

&nbsp; - Calendars.Read

\- Grant admin consent if your tenant requires it



Notes:

\- The OAuth scope requests offline\_access (for refresh tokens). In many tenants this does not need a separate permission entry.



4\) Create a client secret

\- Certificates \& secrets – New client secret

\- Set an expiration you’re comfortable maintaining

\- Create, then copy the secret value immediately (you will not be able to view it again)



5\) Copy values into Hubitat

In Hubitat Calendar Switch app, enter:

\- Microsoft client ID: Application (client) ID

\- Microsoft client secret: the secret value you copied

\- Tenant:

&nbsp; - common is a good default

&nbsp; - alternatives include organizations, consumers, or a specific tenant ID

Then click Authorize Microsoft.



If Microsoft auth fails with “no access token returned”

Common causes:

\- Redirect URI mismatch (must be exactly https://cloud.hubitat.com/oauth/stateredirect)

\- Wrong secret (must be the secret value, not the secret ID)

\- Missing Calendars.Read delegated permission

\- Tenant mismatch (try “common” first, then your tenant GUID if needed)



Using it (create switches and set rules)

Step 1 – Authenticate provider(s)

\- In the app, enter credentials and click Authorize Google and/or Authorize Microsoft.

\- You should see Status: Connected for each provider you authorize.



Step 2 – Add a calendar switch

\- In the app, choose:

&nbsp; - Provider (Google or Microsoft)

&nbsp; - Switch name

&nbsp; - Calendar ID (optional)

&nbsp;   - Google: leave blank for primary, or provide a calendar ID

&nbsp;   - Microsoft: leave blank for primary (future expansion may support selecting other calendars)

\- Click Add switch

\- The switch will appear in the list of switches, with an “Open” link.



Step 3 – Configure rules on the device (recommended flow)

Open the child device and set:

\- Must include words (optional)

&nbsp; - If set, an event must contain at least one word/phrase (comma or new-line separated)

\- Ignore words (optional)

&nbsp; - If any word matches, the event is ignored

\- Minutes before start

\- Minutes after end

\- Filters:

&nbsp; - Only consider busy events

&nbsp; - Allow all-day events

&nbsp; - Allow private events



Step 4 – Test

On the device, use:

\- Apply rules now (uses last fetched events)

\- Fetch now and apply (forces a provider fetch, then evaluates)

\- Test now (fetch + evaluate, and records a result)



Understanding app-level timing vs device-level timing

\- App fetch interval controls how often the app calls Google/Microsoft to get fresh data.

\- The device re-evaluates every minute using the most recently fetched events.

\- If you change device rules and want immediate results, use Fetch now and apply on that device.



Troubleshooting

OAuth is not enabled for this app

\- Go to Apps Code, open the app, enable OAuth, Save, then try again.



Redirect hangs on a blank page

\- This can be normal: some browsers will not auto-close tabs opened during OAuth.

\- If the log shows “Stored token…”, close the tab and return to Hubitat. Refresh the app page and confirm Status: Connected.



Switch does not create / “device type not found”

\- The driver name and namespace must match exactly what the app expects.

\- Confirm you installed the child driver under Drivers Code and that its metadata matches:

&nbsp; - name: Hubitat Calendar Switch – Control Device

&nbsp; - namespace: truittchris



Support

\- Website: https://christruitt.com

\- GitHub: https://github.com/truittchris

\- Tip Jar: https://christruitt.com/tip-jar



When requesting help, include:

\- Hubitat model/firmware

\- App version and driver version

\- Provider (Google or Microsoft)

\- What you expected vs what happened

\- Relevant logs (redact tokens/secrets)



