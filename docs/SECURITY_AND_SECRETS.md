# Security and secrets

This project requires OAuth client credentials for Google and/or Microsoft.

## What is stored on the hub

The app stores:

- OAuth access token
- OAuth refresh token
- Token expiry metadata

These values are stored in app state on your hub.

## What you must keep private

- Google client secret
- Microsoft client secret

Do not:

- post screenshots that include secrets
- paste secrets into GitHub issues or forums
- commit secrets into a repo

## If you accidentally exposed a secret

1. Revoke and rotate the secret in the provider console (Google Cloud Console or Azure Portal).
2. Update the secret in the Hubitat app settings.
3. Disconnect and reconnect the provider in the app.

## Data access scope

Recommended scopes are read-only:

- Google: Calendar read-only
- Microsoft Graph: Calendars.Read (delegated)

Only grant broader scopes if you understand and accept the risk.