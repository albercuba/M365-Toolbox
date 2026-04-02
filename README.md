# M365 Toolbox

M365 Toolbox - Web-based PowerShell operations for Microsoft 365

This project is a Dockerized web application for running approved Microsoft 365 PowerShell scripts from a browser-based interface. It is designed to grow into a reusable operations console with a script catalog, safer parameter handling, and per-run output tracking.

## Current scaffold

- `backend/`
  Express API with a PowerShell runner and script catalog
- `frontend/`
  React/Vite web UI for browsing scripts and launching runs
- `output/`
  Local output folder mounted into the backend container
- `docker-compose.yml`
  Starts the backend and frontend together

## First integrated script

- `M365 - CompromisedAccountRemediation/M365-CompromisedAccountRemediation.ps1`

The backend mounts `../Powershell` into the container at `/workspace-scripts` and runs the approved script from there.

## How it works

1. The frontend loads the script catalog from the backend.
2. The user fills in the approved form fields for a script.
3. The backend converts those values into an allowlisted `pwsh` command.
4. The script runs inside the backend container.
5. Stdout, stderr, status, and generated files are available through the UI and the mounted `output/` folder.

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
- Only allowlisted parameters are exposed by the UI right now.
- The script registry in `backend/src/data/scripts.js` is the place to add future scripts.

## Next growth areas

- authentication and RBAC
- secrets and certificate handling
- run history persistence
- downloadable output browser
- approvals for high-impact scripts
- additional M365 scripts and categories
