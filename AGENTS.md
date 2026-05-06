# AGENTS.md

## Project overview

M365 Toolbox is a web-based console for approved Microsoft 365 PowerShell operations. It combines:

- `frontend/`: React + Vite UI.
- `backend/`: Node.js/Express API, PostgreSQL-backed run store, Redis/BullMQ queue integration, artifact endpoints, and worker health/status APIs.
- `scripts/`: Toolbox-native PowerShell scripts, shared report helpers, and the PowerShell wrapper.
- `scripts/catalog/`: JSON catalog metadata that defines script fields, categories, risks, and runner behavior.
- `output/`: Generated run artifacts. Treat this as runtime output, not source.
- Docker Compose files for local, production, and Coolify/Portainer-style deployments.

The runtime flow is:

```text
Frontend
  -> Express API backend
  -> BullMQ / Redis
  -> PowerShell worker
  -> PostgreSQL run store + output artifacts
```

## General agent instructions

- Keep changes small, focused, and consistent with the existing architecture.
- Prefer editing the correct layer instead of adding workarounds across layers.
- Do not commit generated artifacts from `output/`, logs, secrets, tenant data, migration CSV files, or local environment files.
- Preserve the allowlisted-script execution model. Do not introduce arbitrary command execution from user-controlled frontend input.
- Treat this as an admin/security operations tool. Be conservative with remediation workflows, tenant data, credentials, device-code auth, and generated reports.
- Favor read-only/reporting behavior unless a workflow is explicitly designed as remediation and requires approval.
- Update documentation when changing behavior, environment variables, ports, script catalog metadata, deployment steps, or operator-facing UX.
- When asked for a commit comment/message, use Conventional Commits style and include the type prefix, such as `fix:`, `feat:`, `docs:`, `test:`, `refactor:`, `chore:`, or `ci:`.

## Repository structure

### Backend

Backend source lives in `backend/`.

Common responsibilities:

- Express API routes.
- Script catalog loading.
- Run request validation.
- BullMQ queue integration.
- PostgreSQL persistence through Prisma.
- Artifact listing, signed artifact links, ZIP bundles, and HTML preview support.
- Worker heartbeat and backend status reporting.
- Run cancellation and run lifecycle state handling.

When modifying backend code:

- Keep API validation strict.
- Never trust catalog or frontend input without backend validation.
- Preserve run lifecycle states: queued, running, canceling, completed, failed, canceled, and interrupted.
- Keep artifact paths constrained to configured output directories.
- Do not expose raw filesystem paths, secrets, tokens, passwords, or tenant-sensitive values in API responses.
- Maintain compatibility with the PowerShell worker and wrapper.

### Frontend

Frontend source lives in `frontend/`.

Common responsibilities:

- React/Vite UI.
- Script catalog browsing, search, categories, favorites, and dashboard shortcuts.
- Script parameter forms.
- Device-code sign-in modal.
- Run status polling.
- Run logs and artifact browser.
- Inline HTML report preview.
- Approval confirmation for remediation workflows.
- Backend status UI.

When modifying frontend code:

- Keep operator workflows clear and explicit.
- Preserve approval prompts for remediation or high-impact actions.
- Do not store secrets in local storage.
- Company tenant mappings may be stored client-side, but credentials and migration passwords must not be entered into the browser.
- Keep API calls centralized and consistent with the existing backend endpoints.
- Run the frontend build before finishing UI changes.

### PowerShell scripts

PowerShell scripts live in `scripts/`.

When adding or editing scripts:

- Use toolbox-native scripts only; do not depend on an external script repository mount.
- Prefer report-first/read-only behavior unless the script is explicitly a remediation workflow.
- Redact secrets, passwords, tokens, auth codes, tenant-sensitive values, and migration CSV password fields from logs, diagnostics, reports, and artifact names.
- Use shared helpers such as `Shared-ToolboxReport.ps1` where appropriate.
- Write useful structured output and artifacts for the UI to display.
- Keep scripts compatible with the wrapper `scripts/Invoke-ToolboxScript.ps1`.
- Avoid interactive prompts other than supported Microsoft device-code authentication patterns.
- Make destructive or high-impact actions opt-in, clearly named, and cataloged as remediation/high risk.
- Ensure scripts pass PSScriptAnalyzer with no error-severity findings.

### Script catalog

Catalog files live in `scripts/catalog/`.

When adding or changing catalog entries:

- Ensure the catalog metadata matches the actual PowerShell script parameters.
- Use clear titles, descriptions, categories, and risk labels.
- Mark remediation workflows distinctly from read-only workflows.
- Include only approved input fields.
- Avoid passing arbitrary free-form command fragments.
- Keep tenant/company fields compatible with frontend company mapping resolution.
- Update the README or related docs when adding significant workflows.

### Docker and deployment

Compose files:

- `docker-compose.yml`: local/default deployment.
- `docker-compose.prod.yml`: production-oriented deployment.
- `docker-compose.coolify.yml`: Coolify/Portainer-friendly deployment.

When changing deployment behavior:

- Keep backend, worker, Redis, PostgreSQL, and frontend wiring consistent.
- Preserve the mounted script and output paths:
  - `./scripts` -> `/toolbox-scripts`
  - `./output` -> `/app/output`
- Document new environment variables.
- Ensure production deployments require a strong `ARTIFACT_TOKEN_SECRET`.
- Do not expose backend, Redis, PostgreSQL, or worker services publicly unless the compose variant explicitly intends it.

## Development commands

From the repository root:

```bash
npm install
npm run dev
npm run build
npm run test -w backend
npm run test:frontend
npm run lint
npm run db:migrate
npm run docker:build
npm run docker:up
npm run docker:down
```

Backend-specific commands:

```bash
npm run dev:api -w backend
npm run dev:worker -w backend
npm run test -w backend
npm run db:migrate -w backend
npm run migrate:file-runs -w backend
```

Frontend-specific commands:

```bash
npm run dev -w frontend
npm run build -w frontend
npm run lint -w frontend
npm run test -w frontend
```

Docker:

```bash
docker compose build
docker compose up -d
docker compose logs -f backend
docker compose logs -f worker
docker compose down
```

Production compose:

```bash
docker compose -f docker-compose.prod.yml up -d --build
```

## Validation checklist

Before completing changes, run the most relevant checks.

For backend changes:

```bash
npm run test -w backend
```

For frontend changes:

```bash
npm run build -w frontend
npm run lint -w frontend
npm run test -w frontend
```

For PowerShell script changes:

```powershell
Invoke-ScriptAnalyzer -Path ./scripts -Recurse -Severity Warning,Error
```

For Docker/deployment changes:

```bash
ARTIFACT_TOKEN_SECRET=ci-artifact-token-secret docker compose -f docker-compose.prod.yml build backend worker frontend
```

For dependency changes:

```bash
npm audit --omit=dev -w backend
npm audit --omit=dev -w frontend
```

## Security rules

- Never add hardcoded credentials, tenant IDs, access tokens, refresh tokens, client secrets, passwords, device codes, or private keys.
- Never log IMAP migration passwords or authentication secrets.
- Never accept raw shell, PowerShell, or command-line fragments from frontend input.
- Keep execution constrained to cataloged scripts and approved parameters.
- Preserve artifact signing and short-lived artifact URL behavior.
- Keep CORS restrictive and environment-driven.
- Keep remediation actions behind explicit UI approval.
- Treat generated HTML reports as untrusted output; avoid unsafe rendering changes.
- Keep Gitleaks-compatible secret hygiene.

## Data and artifact handling

- `output/` is for generated reports and run artifacts.
- Do not commit generated reports, customer data, tenant data, mailbox exports, CSV migration files, logs, or ZIP bundles.
- Do not rename artifact conventions without updating backend artifact listing, preview, download, and ZIP behavior.
- Preserve compatibility with existing run history and migration/import behavior unless intentionally changing it.

## Code style

### JavaScript / TypeScript

- Use ES modules.
- Keep backend code simple and explicit.
- Prefer small functions with clear validation boundaries.
- Use existing project patterns before introducing new dependencies.
- Avoid broad catch blocks that hide operational failures.
- Return clear error messages without leaking sensitive values.

### React

- Keep components focused and operator-friendly.
- Avoid introducing global state unless necessary.
- Preserve accessibility for buttons, forms, modals, and status messages.
- Make risky/remediation actions visually and textually explicit.
- Keep run status, queued state, cancellation, and artifact states clear.

### PowerShell

- Use approved verbs where practical.
- Prefer explicit parameters and validation.
- Avoid writing secrets to stdout, stderr, reports, logs, or artifact names.
- Make reports useful for admins: summaries, findings, evidence, and recommended next steps.
- Fail clearly when required Microsoft modules, permissions, or licenses are missing.
- Keep authentication behavior compatible with Microsoft device-code prompts surfaced by the UI.

## Adding a new toolbox script

When adding a new workflow:

1. Add the PowerShell script under `scripts/`.
2. Reuse shared helpers where appropriate.
3. Add or update catalog metadata under `scripts/catalog/`.
4. Ensure backend catalog loading recognizes the entry.
5. Confirm the frontend renders fields correctly.
6. Validate read-only versus remediation labeling.
7. Run PSScriptAnalyzer.
8. Run backend tests and frontend build if API/UI behavior changed.
9. Update README or docs with the new workflow if it is user-visible.

## IMAP migration rules

- Do not ask users to paste mailbox passwords into the browser.
- Migration CSV files should be mounted into the container and referenced by path.
- Redact migration CSV password fields from logs, diagnostics, reports, and artifact names.
- Remember that Microsoft 365 migration services perform mailbox sync after Exchange Online migration endpoint and batch creation.
- Do not implement a custom mailbox copy engine in Node, React, or local PowerShell unless the project direction explicitly changes.

## Pull request expectations

A good PR should include:

- A concise description of the change.
- The affected area: backend, frontend, scripts, catalog, Docker, docs, or CI.
- Any new environment variables or deployment considerations.
- Validation commands run and results.
- Screenshots or notes for meaningful UI changes.
- Security considerations for auth, remediation, artifacts, logging, or tenant data.

## Things agents should avoid

- Do not bypass the catalog and wrapper model.
- Do not introduce direct frontend-to-PowerShell execution.
- Do not make Redis, PostgreSQL, or the worker public.
- Do not weaken remediation approvals.
- Do not store secrets in browser storage.
- Do not commit generated artifacts.
- Do not add large dependencies without a strong reason.
- Do not silently change runtime ports, mount paths, or environment variable semantics.
- Do not remove CI/security checks to make a change pass.
