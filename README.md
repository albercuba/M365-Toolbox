# M365 Toolbox

M365 Toolbox is a Dockerized web console for approved Microsoft 365 PowerShell operations. It provides a script catalog, browser-based execution, device-code sign-in prompts, tracked run history, and inline HTML report previews directly in the UI.

## Project layout

- `backend/`
  Express API, script catalog, PowerShell runner, run tracking, and HTML artifact endpoints
- `frontend/`
  React/Vite UI with script catalog, recent runs, inline report previews, and device-code popup handling
- `scripts/`
  Toolbox-native PowerShell scripts and shared helpers such as `Shared-ToolboxReport.ps1`
- `output/`
  Generated report artifacts mounted into the backend container
- `docker-compose.yml`
  Starts frontend and backend together

## Current script catalog

Core reports and workflows:

- `scripts/M365-CompromisedAccountRemediation.ps1`
- `scripts/M365-MfaReport.ps1`
- `scripts/M365-UsageReport.ps1`
- `scripts/M365-LicensingReport.ps1`
- `scripts/M365-GuestAccessReport.ps1`
- `scripts/M365-ConditionalAccessReport.ps1`
- `scripts/M365-MailForwardingAudit.ps1`
- `scripts/M365-SharedMailboxReport.ps1`
- `scripts/M365-SignInRiskReport.ps1`
- `scripts/M365-TeamsExternalAccessReport.ps1`
- `scripts/M365-SharePointSharingReport.ps1`
- `scripts/M365-SecureScoreSnapshot.ps1`
- `scripts/M365-AdminRoleAudit.ps1`

Additional audit and hygiene reports:

- `scripts/M365-InactiveUsersReport.ps1`
- `scripts/M365-AppConsentAudit.ps1`
- `scripts/M365-MailboxPermissionAudit.ps1`
- `scripts/M365-ExternalSharingLinksReport.ps1`
- `scripts/M365-DistributionGroupAudit.ps1`
- `scripts/M365-ServiceHealthSnapshot.ps1`
- `scripts/M365-AuthenticationPolicyReport.ps1`
- `scripts/M365-PrivilegedAppAudit.ps1`
- `scripts/M365-DkimDmarcReport.ps1`
- `scripts/M365-GroupLifecycleReport.ps1`

## Runtime model

The backend mounts:

- `./scripts` into the container at `/toolbox-scripts`
- `./output` into the container output path used for generated artifacts

Scripts are launched through `scripts/Invoke-ToolboxScript.ps1`, and the backend tracks:

- command line
- stdout
- stderr
- status
- exit code
- generated HTML artifact paths

## UI behavior

- The left script catalog is resizable on desktop.
- Runs can be reopened from `Recent Runs`.
- `Run Details` and `Recent Runs` collapse automatically after successful completion.
- HTML reports are rendered inline beneath `Recent Runs`.
- HTML reports can be downloaded directly from the report card.
- When a script emits the Microsoft device-code prompt, the UI opens a modal with:
  - the device code
  - a clickable link to `https://login.microsoft.com/device`

## Start with Docker

```powershell
cd C:\VSCode\M365-Toolbox
docker compose build
docker compose up -d
```

Then open:

- Frontend: `http://localhost:5173`
- Backend API health: `http://localhost:3001/api/health`

If you changed backend scripts, the script catalog, or PowerShell module installation, rebuild the backend image:

```powershell
docker compose build backend
docker compose up -d backend
```

## Notes

- The backend container installs `pwsh` and the required Microsoft Graph / Exchange Online modules.
- Toolbox-native scripts live in `scripts/` and are exposed through the registry in `backend/src/data/scripts.js`.
- Shared report helpers live in `scripts/Shared-ToolboxReport.ps1`.
- Most HTML-style reports use the shared helper for:
  - device-code Graph auth
  - common HTML dashboard rendering
  - timestamped output naming
  - Graph collection and CSV import helpers
- MFA, usage, and compromised-account workflows remain standalone entry scripts with their own specialized logic.

## How a run works

1. The frontend loads the script catalog from the backend.
2. You choose a script in the left catalog and fill in the allowlisted form fields.
3. The backend turns those values into a controlled `pwsh` invocation.
4. The script runs inside the backend container.
5. The UI polls run status and captures stdout/stderr.
6. If the script produces an HTML dashboard, the backend serves it back to the UI for preview and download.

## Next growth areas

- authentication and RBAC
- run history persistence beyond process memory
- approvals for high-impact workflows
- richer artifact browsing for CSV, XLSX, and logs
- secret and certificate-backed auth options where appropriate
