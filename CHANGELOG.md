# Changelog

All notable changes to Hubitat Calendar Switch are documented in this file.

This project follows Semantic Versioning.

## 1.0.5 - 2026-01-18

Driver

- Reworked driver preferences UX to match the project UX standard
  - Match words and Ignore words live on the Preferences page
  - Added clear buttons on the Preferences page (Clear match words, Clear ignore words, Clear both)
  - Reduced Commands tab to the primary actions (Apply rules now, Fetch now, Fetch now and apply, Test now, Refresh)
  - Added an in-page commandHelp explanation visible on the Commands tab under Current States
- Fixed paragraph rendering by using a top-level preferences block (avoids Script1.paragraph errors)

## 1.0.4 - 2026-01-17

Driver

- Aligned driver name and namespace to the app expectations to ensure child device creation succeeds

## 1.0.3 - 2026-01-17

App

- OAuth reliability fixes for Google and Microsoft
- Improved OAuth completion UX with a success page that attempts to close the authorization tab
- Hardened token exchange and error reporting (more actionable error messages)

## 1.0.2 - 2026-01-17

App

- Reduced OAuth callback URL size and avoided long redirect loops that could trigger HTTP 414 and 431 errors

## 1.0.1 - 2026-01-17

App

- Added provider status display and disconnect actions
- Added initial event polling and on-demand polling from child devices

## 1.0.0 - 2026-01-??

Initial release

- OAuth connection to Google Calendar and Microsoft Graph
- Child switch devices that evaluate events locally
- Provider polling and normalized event delivery
