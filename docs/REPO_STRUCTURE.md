# Repo structure

Recommended layout for this project:

.

- Apps/
  - Hubitat_Calendar_Switch.groovy
- Drivers/
  - Hubitat_Calendar_Switch_Control_Device.groovy
- docs/
  - setup.md
  - troubleshooting.md
  - SECURITY_AND_SECRETS.md
  - REPO_STRUCTURE.md
  - RELEASING.md
- hpm/
  - repository.json
- README.md
- CHANGELOG.md
- LICENSE

Notes

- Apps and Drivers are separated so users can paste code into Hubitat easily and so HPM installs can reference stable raw URLs.
- hpm/repository.json is optional but recommended if you plan to publish via Hubitat Package Manager.
