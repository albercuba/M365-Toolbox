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

Core workflows and reporting:

- `scripts/M365-CompromisedAccountRemediation.ps1`
  Contains and investigates compromised accounts, then generates an incident dashboard.
- `scripts/M365-MfaReport.ps1`
  Builds a tenant-wide MFA coverage report with admin-risk visibility.
- `scripts/M365-UsageReport.ps1`
  Creates OneDrive, SharePoint, and mailbox usage dashboards.
- `scripts/M365-LicensingReport.ps1`
  Reviews subscribed SKUs, assigned licenses, and unlicensed users.
- `scripts/M365-GuestAccessReport.ps1`
  Audits guest accounts, invitation state, stale guests, and external domains.
- `scripts/M365-ConditionalAccessReport.ps1`
  Summarizes Conditional Access policies, state, scope, and grant controls.
- `scripts/M365-MailForwardingAudit.ps1`
  Finds mailbox forwarding and inbox rules that redirect mail.
- `scripts/M365-SharedMailboxReport.ps1`
  Reviews shared mailboxes, forwarding, visibility, and delegate counts.
- `scripts/M365-SignInRiskReport.ps1`
  Shows risky users and detections when Identity Protection data is licensed and available.
- `scripts/M365-TeamsExternalAccessReport.ps1`
  Reviews Teams guest exposure, ownership gaps, and external membership.
- `scripts/M365-SharePointSharingReport.ps1`
  Combines SharePoint tenant sharing settings with site usage inventory.
- `scripts/M365-SecureScoreSnapshot.ps1`
  Captures the latest Secure Score snapshot and top improvement controls.
- `scripts/M365-AdminRoleAudit.ps1`
  Audits privileged role assignments and MFA hygiene for admin accounts.

Audit, hygiene, and operations reports:

- `scripts/M365-InactiveUsersReport.ps1`
  Identifies inactive accounts and likely license-reclaim candidates.
- `scripts/M365-AppConsentAudit.ps1`
  Reviews enterprise app consent grants and high-privilege delegated scopes.
- `scripts/M365-MailboxPermissionAudit.ps1`
  Audits Full Access and Send As delegation across mailboxes.
- `scripts/M365-ExternalSharingLinksReport.ps1`
  Reviews SharePoint external sharing posture and active sites.
- `scripts/M365-DistributionGroupAudit.ps1`
  Audits distribution group ownership, membership, and external sender exposure.
- `scripts/M365-ServiceHealthSnapshot.ps1`
  Captures current Microsoft 365 service health and active advisories.
- `scripts/M365-AuthenticationPolicyReport.ps1`
  Reviews security defaults and authentication methods policy posture.
- `scripts/M365-PrivilegedAppAudit.ps1`
  Inventories service principal credentials and non-human identity exposure.
- `scripts/M365-DkimDmarcReport.ps1`
  Reviews accepted domains for DKIM signing and DMARC presence.
- `scripts/M365-GroupLifecycleReport.ps1`
  Audits Microsoft 365 group ownership, renewal activity, and lifecycle hygiene.

Latest additions:

- `scripts/M365-CAPolicyCoverageReport.ps1`
  Maps Conditional Access inclusion and exclusion coverage across users, groups, guests, and apps.
- `scripts/M365-LegacyAuthExposureReport.ps1`
  Surfaces recent legacy-authentication sign-ins and affected users.
- `scripts/M365-PIMRoleActivationReport.ps1`
  Reviews active and eligible privileged role schedule instances from Entra ID.
- `scripts/M365-DeviceComplianceSnapshot.ps1`
  Summarizes Intune-managed device compliance, ownership, and platform mix.
- `scripts/M365-B2BDirectConnectReport.ps1`
  Reviews cross-tenant access defaults and B2B direct connect posture.
- `scripts/M365-MailTransportRulesAudit.ps1`
  Audits Exchange Online mail transport rules and test/audit mode coverage.
- `scripts/M365-DefenderIncidentSnapshot.ps1`
  Captures current Defender incidents, severity, and status.
- `scripts/M365-PrivilegedGroupAudit.ps1`
  Reviews sensitive groups for owner gaps and membership exposure.
- `scripts/M365-PasswordResetReadinessReport.ps1`
  Estimates self-service password reset readiness from registered auth methods.
- `scripts/M365-OneDriveExternalSharingReport.ps1`
  Reviews OneDrive usage and highlights large or highly active personal sites.

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
- The toolbox now runs only toolbox-native scripts from `scripts/`; there is no external PowerShell repository mount anymore.
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
