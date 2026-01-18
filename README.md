# Hubitat Calendar Switch

Hubitat Calendar Switch connects your Hubitat hub to Google Calendar and Microsoft Outlook (Microsoft 365) using OAuth, then creates child switch devices that turn on and off based on your calendar events.

Design intent

- The app owns OAuth, provider polling, and event retrieval.
- Each child switch owns all match logic (match words, ignore words, timing, filters).

## Features

- OAuth connection to Google and Microsoft
- Creates one child switch per calendar rule
- Normalizes events from each provider and delivers them to each child switch
- Child switch rules:
  - match words and ignore words
  - minutes before start and minutes after end
  - busy only, allow all-day events, allow private events
- Manual controls:
  - Fetch now
  - Fetch now and apply
  - Test now

## Requirements

- Hubitat Elevation hub with Remote Admin / cloud access available for OAuth redirect handling
- A Google Cloud project (for Google Calendar OAuth) and or a Microsoft EnTRA app registration (for Microsoft Graph OAuth)

## Installation

1. In Hubitat, add the app code:
   - Apps Code – New App
   - Paste `Apps/Hubitat_Calendar_Switch.groovy`
   - Save

2. In Hubitat, add the driver code:
   - Drivers Code – New Driver
   - Paste `Drivers/Hubitat_Calendar_Switch_Control_Device.groovy`
   - Save

3. Install the app:
   - Apps – Add User App – Hubitat Calendar Switch

## Provider setup

The redirect URI for both providers is:

`https://cloud.hubitat.com/oauth/stateredirect`

### Google Calendar setup

Summary

- Create an OAuth client (Web application)
- Add the redirect URI above
- Add Calendar read-only scope
- Copy the client ID and client secret into the app

Steps

1. In Google Cloud Console, create or select a project.
2. Configure the OAuth consent screen (External is typical).
3. Create credentials:
   - OAuth client ID
   - Application type: Web application
   - Authorized redirect URIs: `https://cloud.hubitat.com/oauth/stateredirect`
4. In Hubitat Calendar Switch:
   - Enter Google client ID and Google client secret
   - Click Authorize Google
   - Complete the consent flow, then close the tab and return to Hubitat

Notes

- The app requests the scope `https://www.googleapis.com/auth/calendar.readonly`.
- If you change Google project settings, disconnect and re-authorize.

### Microsoft Outlook setup

Summary

- Register an app in Microsoft EnTRA ID (Azure portal)
- Add the redirect URI above
- Add delegated permissions for Calendars.Read and offline_access
- Create a client secret
- Copy the client ID and client secret into the app

Steps

1. In Azure portal, register a new app.
2. Add a redirect URI:
   - Platform: Web
   - Redirect URI: `https://cloud.hubitat.com/oauth/stateredirect`
3. Add API permissions (delegated):
   - Microsoft Graph: Calendars.Read
   - Microsoft Graph: offline_access
4. Create a client secret.
5. In Hubitat Calendar Switch:
   - Enter Microsoft client ID, Microsoft client secret
   - Tenant: usually `common`
   - Click Authorize Microsoft
   - Complete the consent flow, then close the tab and return to Hubitat

Notes

- If your Microsoft environment is tenant-restricted, set Tenant to your tenant ID.

## Creating a calendar switch

1. In the app, confirm your provider shows Status: Connected.
2. Under Add a calendar switch:
   - select Provider
   - enter Switch name
   - optional: Calendar ID (leave blank for primary)
   - click Add switch
3. Open the created device and set:
   - Match words and Ignore words
   - timing and filters
4. Run Test now.

## Child switch usage

The Commands tab includes a built-in help note shown in Current States as `commandHelp`.

Typical flow

- Set match and ignore words in Preferences.
- Use Fetch now and apply to pull events immediately.
- Use Test now to validate your rules.

## Troubleshooting

- If Status is Not connected, verify client ID, client secret, and redirect URI.
- If you see "No access_token returned", check:
  - correct redirect URI
  - correct client secret
  - required permissions and consent
- If switches do not update after changing rules:
  - run Fetch now and apply on the device
  - or wait for the next scheduled poll
- If Add switch does nothing:
  - confirm the driver is installed with the exact name `Hubitat Calendar Switch - Control Device` and namespace `truittchris`

More detail: see `docs/troubleshooting.md`.

## Support

- Website: https://christruitt.com
- GitHub: https://github.com/truittchris
- Tip jar: https://christruitt.com/tip-jar

When requesting help, enable Debug logging and include screenshots of the app main page and the affected device page.

## License

MIT License. See `LICENSE`.
