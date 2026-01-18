# Releasing

This project is typically released as a paired app + driver update.

## Versioning

- Bump the app version when any app logic or UI changes.
- Bump the driver version when any driver UX or evaluation logic changes.
- When releasing a combined update, use the highest of the two versions as the Git tag and release title.

## Release checklist

1) Update versions

- Apps/Hubitat_Calendar_Switch.groovy
  - Update APP_VERSION
- Drivers/Hubitat_Calendar_Switch_Control_Device.groovy
  - Update DRIVER_VERSION

2) Update docs

- CHANGELOG.md
  - Add a new section for the release with the date
- README.md
  - Only update if user-facing behavior or setup steps changed

3) Validate on a hub

- Install updated app and driver code
- Connect Google
- Connect Microsoft
- Create a new switch
- Run Test now and confirm:
  - lastFetch and lastPoll update
  - upcomingSummary updates
  - switch state reflects calendar

4) Tag and publish

- Create a Git tag (example: v1.0.5)
- Create a GitHub release using the matching version
- Paste the relevant changelog section into the release notes

5) Update HPM metadata

- Update hpm/repository.json version and releaseNotes
- Ensure raw file URLs reference the new tag

## Backward compatibility notes

- Keep the child driver name and namespace stable. The app expects the driver name exactly.
- When adding new device attributes, avoid renaming existing attributes to prevent dashboards and rules from breaking.
