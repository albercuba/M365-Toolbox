# M365 Toolbox

M365 Toolbox is a web-based console for approved Microsoft 365 PowerShell operations. It combines a React frontend, an Express API, and a toolbox-native PowerShell script catalog so admins can launch reports and response workflows from the browser, authenticate with Microsoft device code, and review results without leaving the app.

## What the project does

- Presents a categorized script catalog for Microsoft 365 operations
- Runs allowlisted PowerShell scripts through a controlled backend wrapper
- Prompts for Microsoft device-code sign-in when a script requires authentication
- Tracks run status, command details, stdout, stderr, exit code, and artifacts
- Renders HTML reports inline in the UI and supports direct download
- Stores favorites in the browser for faster access to common workflows

The current catalog includes 44 toolbox-native scripts across categories such as Identity, Exchange, Security, SharePoint, Teams, Reporting, Licensing, Devices, Operations, Collaboration, and Incident Response.

## Project layout

- `backend/`
  Express API, script registry, PowerShell runner, run tracking, and HTML artifact endpoints
- `frontend/`
  React/Vite UI with the resizable sidebar, category icons, favorites, recent runs, and inline report preview
- `scripts/`
  Toolbox-native PowerShell scripts, the script wrapper, and shared report helpers
- `output/`
  Generated report artifacts written by backend runs
- `docker-compose.yml`
  Starts the backend and frontend containers together

## Current capabilities

- Resizable left sidebar for the script catalog
- Search and category browsing in the catalog
- Category icons in the sidebar for faster scanning
- Favorites toggle and favorites-only filtering
- Recent run history in the UI
- Inline HTML report preview after successful runs
- Device-code modal that surfaces the sign-in code and login URL
- GitHub repository link in the sidebar footer

## Script catalog highlights

Representative workflows currently included in the toolbox:

- `M365-CompromisedAccountRemediation.ps1`
- `M365-MfaReport.ps1`
- `M365-UsageReport.ps1`
- `M365-LicensingReport.ps1`
- `M365-ConditionalAccessReport.ps1`
- `M365-SecureScoreSnapshot.ps1`
- `M365-AdminRoleAudit.ps1`
- `M365-PrivilegedAppAudit.ps1`
- `M365-DefenderIncidentSnapshot.ps1`
- `M365-DeviceComplianceSnapshot.ps1`
- `M365-BreakGlassAccountAudit.ps1`
- `M365-MailflowConnectorAudit.ps1`

Additional scripts cover guest access, Teams ownership, mailbox permissions, external sharing, service health, mail transport rules, password reset readiness, legacy auth exposure, PIM role activation, app credential expiry, and more.

The registry for the full catalog lives in [backend/src/data/scripts.js](/c:/VSCode/M365-Toolbox/backend/src/data/scripts.js), and the PowerShell entry scripts live in [scripts](/c:/VSCode/M365-Toolbox/scripts).

## Runtime model

The backend loads script metadata from the catalog registry and launches scripts through `scripts/Invoke-ToolboxScript.ps1`. Each run is executed with controlled arguments based on allowlisted form fields.

The backend tracks:

- selected script id
- generated command line
- stdout
- stderr
- run status
- exit code
- timestamps
- HTML artifact paths

By default, Docker mounts:

- `./scripts` to `/toolbox-scripts`
- `./output` to `/app/output`

## UI flow

1. The frontend loads the script catalog from the backend API.
2. You select a script from the sidebar and fill in its approved input fields.
3. The backend transforms those values into a controlled `pwsh` invocation.
4. The script runs and the UI polls for status updates.
5. If a Microsoft device-code prompt appears in stdout, the UI opens a sign-in modal.
6. If the script generates an HTML report, the backend serves it back for inline preview and download.

## Run with Docker

```powershell
cd C:\VSCode\M365-Toolbox
docker compose build
docker compose up -d
```

Open:

- Frontend: `http://localhost:5173`
- Backend health: `http://localhost:3001/api/health`

If you change backend code, script metadata, or PowerShell dependencies, rebuild the backend image:

```powershell
docker compose build backend
docker compose up -d backend
```

## Run locally for development

Install workspace dependencies:

```powershell
cd C:\VSCode\M365-Toolbox
npm install
```

Start both apps together:

```powershell
npm run dev
```

That starts:

- backend on `http://localhost:3001`
- frontend dev server on `http://localhost:5173`

The frontend Vite config proxies `/api` requests to the backend during local development.

## Notes

- The backend uses in-memory run tracking today, so history resets when the backend restarts.
- Toolbox-native scripts are served only from `scripts/`; there is no external PowerShell repository mount.
- Shared helpers such as `Shared-ToolboxReport.ps1` support common HTML dashboard rendering and output handling.
- The backend API currently exposes script listing, run creation, run status, and HTML artifact retrieval endpoints.
- CORS is restricted to configured origins, localhost, and private IPv4 ranges.

## Next growth areas

- authentication and RBAC
- persistent run history storage
- approvals for high-impact workflows
- richer artifact browsing for CSV, XLSX, and logs
- localization when multilingual support becomes a priority
- alternative auth flows such as certificate-based or app-based execution where appropriate
