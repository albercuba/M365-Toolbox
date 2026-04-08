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
- `scripts/M365-MFARegistrationCampaignReport.ps1`
  Reviews MFA registration campaign settings and rollout posture.
- `scripts/M365-EnterpriseAppsInventory.ps1`
  Inventories enterprise applications for ownership, visibility, and risk review.
- `scripts/M365-GuestInvitationFailures.ps1`
  Surfaces guest invitation issues and onboarding failures for follow-up.
- `scripts/M365-MailboxAutoReplyAudit.ps1`
  Reviews mailbox automatic replies for external exposure and operational awareness.
- `scripts/M365-CalendarSharingAudit.ps1`
  Audits mailbox calendar sharing posture and external visibility.
- `scripts/M365-RoleEligibleAssignmentsReport.ps1`
  Reviews eligible privileged role assignments and standing access exposure.
- `scripts/M365-AnonymousLinkExposureReport.ps1`
  Highlights anonymous sharing links and externally accessible content exposure.
- `scripts/M365-TeamsOwnershipAudit.ps1`
  Reviews Microsoft Teams ownership gaps and governance hygiene.
- `scripts/M365-AppCredentialExpiryReport.ps1`
  Finds expiring or stale application credentials that need rotation.
- `scripts/M365-MailflowConnectorAudit.ps1`
  Reviews inbound and outbound Exchange Online connectors, relay paths, and TLS posture.
- `scripts/M365-BreakGlassAccountAudit.ps1`
  Reviews emergency access accounts for privilege, MFA, and recent sign-in visibility.

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
