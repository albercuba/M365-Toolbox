# M365 Toolbox

M365 Toolbox - Web-based PowerShell operations for Microsoft 365

This project is a Dockerized web application for running approved Microsoft 365 PowerShell scripts from a browser-based interface. It is designed to grow into a reusable operations console with a script catalog, safer parameter handling, per-run output tracking, and report previews directly in the UI.

## Current scaffold

- `backend/`
  Express API with a PowerShell runner and script catalog
- `frontend/`
  React/Vite web UI for browsing scripts and launching runs
- `output/`
  Local output folder mounted into the backend container
- `docker-compose.yml`
  Starts the backend and frontend together

## Integrated scripts

- `M365 - CompromisedAccountRemediation/M365-CompromisedAccountRemediation.ps1`
- `scripts/M365-MfaReport.ps1`

The backend mounts:

- `../Powershell` into the container at `/workspace-scripts`
- `./scripts` into the container at `/toolbox-scripts`

The compromised-account workflow runs from `/workspace-scripts`, while toolbox-native scripts such as the MFA report run from `/toolbox-scripts`.

## How it works

1. The frontend loads the script catalog from the backend.
2. The user fills in the approved form fields for a script.
3. The backend converts those values into an allowlisted `pwsh` command.
4. The script runs inside the backend container.
5. Stdout, stderr, status, and generated files are available through the UI and the mounted `output/` folder.
6. For the MFA report, the UI detects the Microsoft device-code prompt, shows the code in a popup, and lets you open the Microsoft device login page directly.
7. Completed MFA runs can render their generated HTML dashboard inline and offer a direct HTML download.

## Start with Docker

```powershell
cd C:\VSCode\M365-Toolbox
docker compose build
docker compose up
```

Then open:

- Frontend: `http://localhost:5173`
- Backend API: `http://localhost:3001/api/health`

## Notes

- The backend container installs `pwsh` so it can run PowerShell scripts directly.
- The mounted `../Powershell` folder is read-only inside the container.
- The mounted `./scripts` folder is also read-only inside the container and is intended for toolbox-managed scripts.
- Only allowlisted parameters are exposed by the UI right now.
- The script registry in `backend/src/data/scripts.js` is the place to add future scripts.
- The MFA report uses Microsoft Graph device-code authentication inside the web UI and browser-auth/device-auth behavior can differ between local PowerShell and the containerized runner.

## MFA report flow

1. Open `M365 MFA Report` in the script catalog.
2. Start the run from the web UI.
3. When the script prints the Microsoft device-code prompt, the UI opens a popup with:
   - the device code
   - a clickable link to `https://login.microsoft.com/device`
4. Complete sign-in with an admin account that has the required Microsoft Graph permissions.
5. After the run finishes, the generated HTML dashboard appears in the UI under the report card and can be downloaded directly.

## Next growth areas

- authentication and RBAC
- secrets and certificate handling
- run history persistence
- downloadable output browser beyond the MFA HTML report
- approvals for high-impact scripts
- additional M365 scripts and categories
