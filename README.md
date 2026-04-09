# M365 Toolbox

M365 Toolbox is a web-based console for approved Microsoft 365 PowerShell operations. It combines a React frontend, an Express API, and a toolbox-native PowerShell script catalog so admins can launch reports and response workflows from the browser, authenticate with Microsoft device code, and review results without leaving the app.

## What the project does

- Presents a categorized script catalog for Microsoft 365 operations
- Runs allowlisted PowerShell scripts through a controlled backend wrapper
- Prompts for Microsoft device-code sign-in when a script requires authentication
- Tracks run status, command details, stdout, stderr, exit code, and artifacts
- Renders HTML reports inline in the UI and supports direct download
- Stores favorites in the browser for faster access to common workflows

The current catalog includes 57 toolbox-native scripts across categories such as Identity, Exchange, Security, SharePoint, Teams, Reporting, Licensing, Devices, Operations, Collaboration, and Incident Response.

## Project layout

- `backend/`
  Express API, script registry, PowerShell runner, run tracking, and HTML artifact endpoints
- `frontend/`
  React/Vite UI with the resizable sidebar, category icons, dark mode, favorites, dashboard shortcuts, run details, and inline report preview
- `scripts/`
  Toolbox-native PowerShell scripts, the script wrapper, and shared report helpers
- `output/`
  Generated report artifacts written by backend runs
- `docker-compose.yml`
  Starts the backend and frontend containers together
- `docker-compose.prod.yml`
  Production-oriented Docker Compose deployment for self-hosting
- `docker-compose.coolify.yml`
  Compose variant aimed at Coolify or Portainer-style deployments

## Current capabilities

- Resizable left sidebar for the script catalog
- Search and category browsing in the catalog
- Category icons in the sidebar for faster scanning
- Read-only versus remediation mode labeling and filtering
- Favorites toggle and favorites-only filtering
- Dark mode toggle
- Better dashboard home screen with quick stats and shortcuts
- Recent favorites and most-used shortcuts on the dashboard
- Persistent run history in the UI
- Queue-aware execution with queued run status
- Run cancellation from the UI
- Structured run logs, timestamps, and clearer run-state feedback
- Artifact browser for HTML, CSV, XLSX, text, and log downloads
- Inline HTML report preview after successful runs, with the report brought into focus automatically
- Device-code modal that surfaces the sign-in code, copy action, and login URL
- Approval confirmation prompt for remediation workflows
- Action profiles and high-impact action guidance for the compromised-account workflow
- Backend status view for PowerShell, modules, script mount, and output path readiness
- GitHub repository link in the sidebar footer

## Current script catalog

Core workflows and reporting:

- `scripts/M365-CompromisedAccountRemediation.ps1`
  Contains and investigates compromised accounts with review-only and containment profiles, recent sign-in review, optional delegate cleanup and protocol lock-down, and a richer incident dashboard.
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
- `scripts/M365-ImpossibleTravelReview.ps1`
  Highlights rapid country-to-country sign-in changes for the same user.
- `scripts/M365-MailboxLoginAnomalyReview.ps1`
  Reviews mailbox-related sign-ins with suspicious client, failure, or access patterns.
- `scripts/M365-ExternalTenantTrustReview.ps1`
  Summarizes cross-tenant partner trust settings and guest domain concentration.
- `scripts/M365-PrivilegedUserSignInReview.ps1`
  Reviews recent sign-in visibility for privileged users across directory roles.
- `scripts/M365-InboxRuleThreatHunt.ps1`
  Hunts for suspicious inbox rules that forward, redirect, hide, or delete messages.
- `scripts/M365-ConditionalAccessGapReview.ps1`
  Highlights Conditional Access policies with exclusions, missing controls, or weak enforcement.
- `scripts/M365-MFAExclusionAudit.ps1`
  Audits MFA-enforcing policies that still exclude users, groups, or roles.
- `scripts/M365-DormantAdminAccountReview.ps1`
  Finds privileged accounts with little or no recent sign-in activity.
- `scripts/M365-SharedMailboxAbuseReview.ps1`
  Reviews shared mailboxes for forwarding, risky delegates, and inbox rule abuse signals.
- `scripts/M365-TransportRuleThreatHunt.ps1`
  Flags transport rules with redirect, bypass, delete, or routing threat signals.
- `scripts/M365-BreakGlassAccountHardeningReview.ps1`
  Reviews likely break-glass accounts for MFA, state, and hardening posture.
- `scripts/M365-OAuthAppRiskReview.ps1`
  Highlights OAuth-enabled apps with delegated grants, risky scopes, and credential risk.

The registry for the full catalog lives in [backend/src/data/scripts.js](/c:/VSCode/M365-Toolbox/backend/src/data/scripts.js), and the PowerShell entry scripts live in [scripts](/c:/VSCode/M365-Toolbox/scripts).

## Runtime model

The backend loads script metadata from the catalog registry and launches scripts through `scripts/Invoke-ToolboxScript.ps1`. Each run is executed with controlled arguments based on allowlisted form fields.

The backend tracks:

- selected script id
- generated command line
- stdout
- stderr
- structured log entries
- run status
- queued, running, canceling, completed, failed, canceled, and interrupted lifecycle states
- exit code
- timestamps
- artifact inventory for exported files

Runtime controls available through environment variables:

- `MAX_CONCURRENT_RUNS`
  Limits how many PowerShell runs can execute at the same time. Extra runs stay queued until a slot opens.
- `RUN_RETENTION_HOURS`
  Controls how long persisted run records are kept before cleanup removes them.
- `RUN_STATE_DIR`
  Overrides the directory used for persisted run state files.
- `OUTPUT_DIR`
  Controls where generated artifacts and backend run state are written.
- `TOOLBOX_SCRIPT_MOUNT_ROOT`
  Controls where the backend resolves toolbox PowerShell scripts.

By default, Docker mounts:

- `./scripts` to `/toolbox-scripts`
- `./output` to `/app/output`

## UI flow

1. The frontend loads the script catalog from the backend API.
2. You select a script from the sidebar and fill in its approved input fields.
3. The backend transforms those values into a controlled `pwsh` invocation.
4. The script runs and the UI polls for status updates.
5. If the backend concurrency limit is busy, the run stays queued until a slot opens.
6. If a remediation workflow is selected, the UI requires explicit approval confirmation before launch.
7. If a Microsoft device-code prompt appears in stdout or stderr, the UI opens a sign-in modal with the code, a copy button, and a direct login link.
8. If the script generates artifacts, the backend serves them back for browsing, preview, and download.
9. If an HTML report is generated, the UI collapses the other run cards and scrolls the report preview into view.

## Deploy with Docker Compose

Docker Compose is the easiest way to deploy the project because it builds both services, wires the frontend to the backend, and mounts the folders used for scripts and generated output.

What Compose starts:

- `backend`
  Node/Express API with `pwsh`, Microsoft Graph modules, Exchange Online Management, and the script runner
- `frontend`
  Nginx container serving the built React app

What Compose mounts:

- `./scripts` into the backend container at `/toolbox-scripts` as read-only
- `./output` into the backend container at `/app/output` for generated artifacts

What Compose exposes:

- `http://localhost:5173` for the frontend
- `http://localhost:3001` for the backend API

Available compose files:

- `docker-compose.yml`
  Local default compose file that publishes the frontend on port `5173`
- `docker-compose.prod.yml`
  Production-oriented self-hosted deployment that publishes the frontend on port `8080` by default
- `docker-compose.coolify.yml`
  Platform-friendly variant that uses a named Docker volume for report output and keeps the backend internal

Default deployment steps:

```powershell
cd C:\VSCode\M365-Toolbox
docker compose build
docker compose up -d
```

After deployment:

- Frontend: `http://localhost:5173`
- Backend health: `http://localhost:3001/api/health`

Useful Docker Compose commands:

```powershell
docker compose ps
docker compose logs -f
docker compose logs -f backend
docker compose logs -f frontend
docker compose down
```

Deploy with the production compose file:

```powershell
cd C:\VSCode\M365-Toolbox
docker compose -f docker-compose.prod.yml up -d --build
```

Default production ports:

- Frontend: `http://localhost:8080`
- Backend health: `http://localhost:3001/api/health`

The production compose file supports these environment overrides:

- `FRONTEND_ORIGIN`
  Public frontend URL used by backend CORS validation
- `FRONTEND_PORT`
  Host port mapped to the frontend container, default `8080`
- `BACKEND_PORT`
  Host port mapped to the backend container, default `3001`

Example:

```powershell
$env:FRONTEND_ORIGIN="https://toolbox.example.com"
$env:FRONTEND_PORT="80"
$env:BACKEND_PORT="3001"
docker compose -f docker-compose.prod.yml up -d --build
```

Deploy with the Coolify or Portainer-friendly compose file:

```powershell
docker compose -f docker-compose.coolify.yml up -d --build
```

Why this variant is different:

- the backend uses `expose` instead of a public host port
- report output is stored in a named Docker volume called `toolbox_output`
- runtime values are expected to be supplied by the platform UI

Recommended variables for Coolify or Portainer:

- `FRONTEND_ORIGIN`
  Set this to the public URL of the deployed frontend, for example `https://toolbox.example.com`
- `FRONTEND_PORT`
  Optional host port override if you are not using a platform-managed proxy

When to rebuild:

- Rebuild `backend` if you change backend code, script metadata, PowerShell scripts, or PowerShell module installation
- Rebuild `frontend` if you change the React UI and want the containerized build updated
- Rebuild everything if you change shared package metadata or Dockerfiles

Examples:

```powershell
docker compose build backend
docker compose up -d backend
```

```powershell
docker compose build frontend
docker compose up -d frontend
```

Notes for deployment:

- The backend persists run history to a state file, so completed runs remain visible after backend restarts.
- Generated artifacts remain in the host `output/` folder because it is mounted into the container.
- Scripts are loaded from the host `scripts/` folder, so keep that directory in place on the machine running Docker Compose.
- The Coolify or Portainer-friendly compose file stores artifacts in a Docker-managed volume instead of the host `output/` folder.

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

## Implementation notes

- Run history is persisted to a backend state file so completed runs survive backend restarts.
- Running or queued jobs are marked as interrupted if the backend restarts before they finish.
- Toolbox-native scripts are served only from `scripts/`; there is no external PowerShell repository mount.
- Shared helpers such as `Shared-ToolboxReport.ps1` support common HTML dashboard rendering and output handling.
- The backend API exposes script listing, run creation, run status, cancellation, artifact listing, artifact download, HTML preview, and backend status endpoints.
- Input values are validated on the backend before PowerShell execution starts.
- Run retention is controlled by backend retention settings so old run records do not accumulate forever.
- CORS is restricted to configured origins, localhost, and private IPv4 ranges.

## Product direction

The current implementation focuses on safe execution, persistent run visibility, and better operator feedback. The next natural product steps are:

- authentication and RBAC for multi-user environments
- a real database-backed run store instead of file-based persistence
- richer approval workflows with requester and approver identities
- stronger artifact management and retention controls
- localization when multilingual support becomes a priority
- alternative auth flows such as certificate-based or app-based execution where appropriate
