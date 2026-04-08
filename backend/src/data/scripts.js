const tenantField = {
  id: "tenantId",
  label: "Tenant ID or Domain",
  type: "text",
  required: false,
  placeholder: "contoso.onmicrosoft.com",
  helpText: "Optional. If provided, device-code sign-in is scoped to this tenant."
};

const compromisedAccountScript = {
  id: "m365-compromised-account-remediation",
  name: "M365 Compromised Account Remediation",
  category: "Incident Response",
  summary: "Contain a compromised Microsoft 365 account and generate an incident dashboard.",
  description:
    "Runs the approved compromised-account workflow against one or more target users, authenticates by device code, performs the selected review or remediation actions, and exports an HTML dashboard into the toolbox output folder.",
  scriptRelativePath: "M365-CompromisedAccountRemediation.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  outputs:
    "Writes an HTML incident dashboard plus supporting CSV and log artifacts to the configured output directory.",
  fields: [
    {
      id: "userPrincipalName",
      label: "User Principal Names",
      type: "textarea",
      required: true,
      placeholder: "user1@domain.com, user2@domain.com",
      helpText: "Comma or newline separated UPNs."
    },
    {
      id: "actions",
      label: "Actions",
      type: "multiselect",
      required: true,
      defaultValue: [
        "DisableUser",
        "RevokeSessions",
        "ResetPassword",
        "RemoveMfaMethods",
        "DisableInboxRules",
        "RemoveMailboxForwarding",
        "DisableSignature",
        "ExportAuditLog"
      ],
      options: [
        "DisableUser",
        "RevokeSessions",
        "ResetPassword",
        "ReviewMfaMethods",
        "RemoveMfaMethods",
        "DisableInboxRules",
        "ReviewMailboxForwarding",
        "RemoveMailboxForwarding",
        "DisableSignature",
        "ReviewMailboxDelegates",
        "ExportAuditLog"
      ]
    },
    {
      id: "auditLogDays",
      label: "Audit Window (days)",
      type: "number",
      required: false,
      defaultValue: 10,
      min: 1,
      max: 30
    },
    tenantField,
    {
      id: "includeGeneratedPasswordsInResults",
      label: "Include generated passwords in results",
      type: "checkbox",
      defaultValue: false
    },
    {
      id: "whatIf",
      label: "Preview only (WhatIf)",
      type: "checkbox",
      defaultValue: true
    }
  ]
};

const checkMfaStatusScript = {
  id: "m365-check-mfa-status",
  name: "M365 MFA Report",
  category: "Identity",
  summary: "Generate a full tenant MFA report with coverage, admin risk, and exportable dashboards.",
  description:
    "Runs the approved Microsoft 365 MFA reporting workflow, prompts for admin sign-in by device code, and exports Excel and HTML reports into the toolbox output folder.",
  scriptRelativePath: "M365-MfaReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  outputs:
    "Writes Excel and HTML MFA reports to the configured output directory.",
  fields: [
    {
      id: "includeGuests",
      label: "Include guest users",
      type: "checkbox",
      defaultValue: false,
      helpText: "After the run starts, use the device code shown in the output to sign in with an admin account."
    },
    tenantField
  ]
};

const usageReportScript = {
  id: "m365-usage-report",
  name: "M365 Usage Report",
  category: "Reporting",
  summary: "Generate OneDrive, SharePoint, and Mailbox storage usage dashboards.",
  description:
    "Runs the approved Microsoft 365 storage usage workflow with Microsoft Graph device-code authentication and exports an HTML dashboard into the toolbox output folder.",
  scriptRelativePath: "M365-UsageReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  outputs:
    "Writes an HTML dashboard with OneDrive, SharePoint, and mailbox usage sections to the configured output directory.",
  fields: [
    {
      id: "reports",
      label: "Reports",
      type: "multiselect",
      required: true,
      defaultValue: ["OneDrive", "SharePoint", "Mailbox"],
      options: ["OneDrive", "SharePoint", "Mailbox"],
      helpText: "Select one or more report types to export."
    },
    tenantField
  ]
};
const licensingReportScript = {
  id: "m365-licensing-report",
  name: "M365 Licensing Report",
  category: "Licensing",
  summary: "Review assigned licenses, unlicensed users, and SKU consumption across the tenant.",
  description:
    "Collects tenant SKU inventory and user license assignments, then exports a light HTML dashboard with SKU consumption, unlicensed users, and multi-license visibility.",
  scriptRelativePath: "M365-LicensingReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-licensing-report",
  outputs: "Writes an HTML licensing dashboard to the configured output directory.",
  fields: [
    {
      id: "includeDisabledUsers",
      label: "Include disabled users",
      type: "checkbox",
      defaultValue: false
    },
    tenantField
  ]
};

const guestAccessReportScript = {
  id: "m365-guest-access-report",
  name: "M365 Guest Access Report",
  category: "Identity",
  summary: "Review guest user inventory, stale guest accounts, and external tenant domains.",
  description:
    "Collects guest identities from Microsoft Entra ID and exports an HTML dashboard covering invitation state, stale guest access, and domain concentration.",
  scriptRelativePath: "M365-GuestAccessReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-guest-access-report",
  outputs: "Writes an HTML guest access dashboard to the configured output directory.",
  fields: [
    {
      id: "staleDays",
      label: "Stale guest threshold (days)",
      type: "number",
      required: false,
      defaultValue: 90,
      min: 30,
      max: 365
    },
    tenantField
  ]
};

const conditionalAccessReportScript = {
  id: "m365-conditional-access-report",
  name: "M365 Conditional Access Report",
  category: "Identity",
  summary: "Inventory Conditional Access policies with state, scope, and grant controls.",
  description:
    "Exports a Conditional Access dashboard that highlights enabled, report-only, and disabled policies along with their included scope and grant configuration.",
  scriptRelativePath: "M365-ConditionalAccessReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-conditional-access-report",
  outputs: "Writes an HTML Conditional Access dashboard to the configured output directory.",
  fields: [
    {
      id: "includeDisabledPolicies",
      label: "Include disabled policies",
      type: "checkbox",
      defaultValue: true
    },
    tenantField
  ]
};

const mailForwardingAuditScript = {
  id: "m365-mail-forwarding-audit",
  name: "M365 Mail Forwarding Audit",
  category: "Exchange",
  summary: "Audit mailbox forwarding settings and inbox rules that redirect or forward mail.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard with mailbox forwarding targets and inbox rules that send mail elsewhere.",
  scriptRelativePath: "M365-MailForwardingAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-mail-forwarding-audit",
  outputs: "Writes an HTML mail forwarding dashboard to the configured output directory.",
  fields: [
    {
      id: "includeInboxRules",
      label: "Inspect inbox rules",
      type: "checkbox",
      defaultValue: true
    },
    tenantField
  ]
};

const sharedMailboxReportScript = {
  id: "m365-shared-mailbox-report",
  name: "M365 Shared Mailbox Report",
  category: "Exchange",
  summary: "Review shared mailbox inventory, forwarding, visibility, and delegate coverage.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard covering shared mailbox inventory and optional delegate counts.",
  scriptRelativePath: "M365-SharedMailboxReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-shared-mailbox-report",
  outputs: "Writes an HTML shared mailbox dashboard to the configured output directory.",
  fields: [
    {
      id: "includePermissions",
      label: "Collect delegate counts",
      type: "checkbox",
      defaultValue: true
    },
    tenantField
  ]
};

const signInRiskReportScript = {
  id: "m365-sign-in-risk-report",
  name: "M365 Sign-In Risk Report",
  category: "Identity",
  summary: "Review risky users and risk detections from Microsoft Entra ID.",
  description:
    "Exports an HTML risk dashboard with current risky identities and recent risk detections within the selected lookback window.",
  scriptRelativePath: "M365-SignInRiskReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-sign-in-risk-report",
  outputs: "Writes an HTML sign-in risk dashboard to the configured output directory.",
  fields: [
    {
      id: "lookbackDays",
      label: "Lookback window (days)",
      type: "number",
      required: false,
      defaultValue: 30,
      min: 7,
      max: 180
    },
    tenantField
  ]
};

const teamsExternalAccessReportScript = {
  id: "m365-teams-external-access-report",
  name: "M365 Teams External Access Report",
  category: "Teams",
  summary: "Review Teams guest exposure, team ownership, and external member concentration.",
  description:
    "Exports an HTML dashboard that inspects Teams-backed Microsoft 365 groups and highlights guest membership exposure and ownership gaps.",
  scriptRelativePath: "M365-TeamsExternalAccessReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-teams-external-access-report",
  outputs: "Writes an HTML Teams exposure dashboard to the configured output directory.",
  fields: [
    {
      id: "maxTeamsToInspect",
      label: "Maximum teams to inspect",
      type: "number",
      required: false,
      defaultValue: 100,
      min: 10,
      max: 500
    },
    tenantField
  ]
};

const sharePointSharingReportScript = {
  id: "m365-sharepoint-sharing-report",
  name: "M365 SharePoint Sharing Report",
  category: "SharePoint",
  summary: "Review SharePoint tenant sharing posture and active site inventory.",
  description:
    "Exports a SharePoint dashboard that combines tenant sharing settings with site usage inventory for the selected reporting period.",
  scriptRelativePath: "M365-SharePointSharingReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-sharepoint-sharing-report",
  outputs: "Writes an HTML SharePoint sharing dashboard to the configured output directory.",
  fields: [
    {
      id: "reportPeriod",
      label: "Usage period",
      type: "text",
      required: false,
      defaultValue: "D30",
      placeholder: "D7, D30, D90, D180",
      helpText: "Accepted values: D7, D30, D90, D180."
    },
    tenantField
  ]
};

const secureScoreSnapshotScript = {
  id: "m365-secure-score-snapshot",
  name: "M365 Secure Score Snapshot",
  category: "Security",
  summary: "Capture the latest Secure Score snapshot and top improvement actions.",
  description:
    "Exports an HTML Secure Score dashboard with the latest score snapshot and the highest-value improvement controls returned by Microsoft Graph.",
  scriptRelativePath: "M365-SecureScoreSnapshot.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-secure-score-snapshot",
  outputs: "Writes an HTML Secure Score dashboard to the configured output directory.",
  fields: [
    {
      id: "topActions",
      label: "Top controls to include",
      type: "number",
      required: false,
      defaultValue: 15,
      min: 5,
      max: 50
    },
    tenantField
  ]
};

const adminRoleAuditScript = {
  id: "m365-admin-role-audit",
  name: "M365 Admin Role Audit",
  category: "Identity",
  summary: "Review privileged role assignments and MFA hygiene for admin accounts.",
  description:
    "Exports an HTML audit dashboard for privileged directory roles, assigned members, and whether those privileged accounts have MFA methods registered.",
  scriptRelativePath: "M365-AdminRoleAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-admin-role-audit",
  outputs: "Writes an HTML admin role audit dashboard to the configured output directory.",
  fields: [tenantField]
};

const inactiveUsersReportScript = {
  id: "m365-inactive-users-report",
  name: "M365 Inactive Users Report",
  category: "Identity",
  summary: "Find inactive accounts and identify licensed users with no recent sign-in activity.",
  description:
    "Exports an HTML dashboard that highlights inactive accounts, last sign-in dates, and likely license reclaim opportunities.",
  scriptRelativePath: "M365-InactiveUsersReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-inactive-users-report",
  outputs: "Writes an HTML inactive-user dashboard to the configured output directory.",
  fields: [
    {
      id: "inactiveDays",
      label: "Inactive threshold (days)",
      type: "number",
      required: false,
      defaultValue: 90,
      min: 30,
      max: 365
    },
    {
      id: "includeDisabledUsers",
      label: "Include disabled users",
      type: "checkbox",
      defaultValue: false
    },
    tenantField
  ]
};

const appConsentAuditScript = {
  id: "m365-app-consent-audit",
  name: "M365 App Consent Audit",
  category: "Security",
  summary: "Review enterprise app consents and delegated permission grants.",
  description:
    "Exports an HTML dashboard for OAuth delegated grants, high-privilege scopes, and consent exposure across enterprise apps.",
  scriptRelativePath: "M365-AppConsentAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-app-consent-audit",
  outputs: "Writes an HTML app consent dashboard to the configured output directory.",
  fields: [tenantField]
};

const mailboxPermissionAuditScript = {
  id: "m365-mailbox-permission-audit",
  name: "M365 Mailbox Permission Audit",
  category: "Exchange",
  summary: "Review mailbox Full Access and Send As delegation across user and shared mailboxes.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard that summarizes mailbox permission exposure.",
  scriptRelativePath: "M365-MailboxPermissionAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-mailbox-permission-audit",
  outputs: "Writes an HTML mailbox permission dashboard to the configured output directory.",
  fields: [tenantField]
};

const externalSharingLinksReportScript = {
  id: "m365-external-sharing-links-report",
  name: "M365 External Sharing Links Report",
  category: "SharePoint",
  summary: "Review SharePoint external sharing posture and active sites with shared-content exposure.",
  description:
    "Exports an HTML dashboard that combines SharePoint sharing settings with active site inventory for the selected usage period.",
  scriptRelativePath: "M365-ExternalSharingLinksReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-external-sharing-links-report",
  outputs: "Writes an HTML external-sharing dashboard to the configured output directory.",
  fields: [
    {
      id: "reportPeriod",
      label: "Usage period",
      type: "text",
      required: false,
      defaultValue: "D30",
      placeholder: "D7, D30, D90, D180",
      helpText: "Accepted values: D7, D30, D90, D180."
    },
    tenantField
  ]
};

const distributionGroupAuditScript = {
  id: "m365-distribution-group-audit",
  name: "M365 Distribution Group Audit",
  category: "Exchange",
  summary: "Review distribution group ownership, member counts, and exposure to external senders.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard for distribution group hygiene and sender scope.",
  scriptRelativePath: "M365-DistributionGroupAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-distribution-group-audit",
  outputs: "Writes an HTML distribution-group dashboard to the configured output directory.",
  fields: [tenantField]
};

const serviceHealthSnapshotScript = {
  id: "m365-service-health-snapshot",
  name: "M365 Service Health Snapshot",
  category: "Operations",
  summary: "Capture current Microsoft 365 service health and active advisories.",
  description:
    "Exports an HTML dashboard with current service health overviews and active incident/advisory records from Microsoft 365.",
  scriptRelativePath: "M365-ServiceHealthSnapshot.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-service-health-snapshot",
  outputs: "Writes an HTML service-health dashboard to the configured output directory.",
  fields: [tenantField]
};

const authenticationPolicyReportScript = {
  id: "m365-authentication-policy-report",
  name: "M365 Authentication Policy Report",
  category: "Identity",
  summary: "Review security defaults and Microsoft Entra authentication methods policy posture.",
  description:
    "Exports an HTML dashboard for security defaults and authentication methods policy settings returned by Microsoft Graph.",
  scriptRelativePath: "M365-AuthenticationPolicyReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-authentication-policy-report",
  outputs: "Writes an HTML authentication-policy dashboard to the configured output directory.",
  fields: [tenantField]
};

const privilegedAppAuditScript = {
  id: "m365-privileged-app-audit",
  name: "M365 Privileged App Audit",
  category: "Security",
  summary: "Review enterprise apps with secrets and certificates used by non-human identities.",
  description:
    "Exports an HTML dashboard that inventories service-principal credentials and highlights apps with elevated credential presence.",
  scriptRelativePath: "M365-PrivilegedAppAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-privileged-app-audit",
  outputs: "Writes an HTML privileged-app dashboard to the configured output directory.",
  fields: [tenantField]
};

const dkimDmarcReportScript = {
  id: "m365-dkim-dmarc-report",
  name: "M365 DKIM / DMARC Report",
  category: "Exchange",
  summary: "Review accepted domains for DKIM signing state and DMARC visibility.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard with DKIM state and DMARC DNS visibility for accepted domains.",
  scriptRelativePath: "M365-DkimDmarcReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-dkim-dmarc-report",
  outputs: "Writes an HTML DKIM/DMARC dashboard to the configured output directory.",
  fields: [tenantField]
};

const groupLifecycleReportScript = {
  id: "m365-group-lifecycle-report",
  name: "M365 Group Lifecycle Report",
  category: "Collaboration",
  summary: "Review Microsoft 365 group ownership, renewal activity, and lifecycle hygiene.",
  description:
    "Exports an HTML dashboard for Unified group ownership gaps, membership size, and renewal status across the selected group sample.",
  scriptRelativePath: "M365-GroupLifecycleReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-group-lifecycle-report",
  outputs: "Writes an HTML group-lifecycle dashboard to the configured output directory.",
  fields: [
    {
      id: "maxGroupsToInspect",
      label: "Maximum groups to inspect",
      type: "number",
      required: false,
      defaultValue: 200,
      min: 50,
      max: 500
    },
    tenantField
  ]
};

const caPolicyCoverageReportScript = {
  id: "m365-ca-policy-coverage-report",
  name: "M365 CA Policy Coverage Report",
  category: "Identity",
  summary: "Map Conditional Access inclusion and exclusion scope across users, groups, guests, and apps.",
  description:
    "Exports an HTML dashboard showing Conditional Access policy scope, exclusions, guest coverage, and app targeting.",
  scriptRelativePath: "M365-CAPolicyCoverageReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-ca-policy-coverage-report",
  outputs: "Writes an HTML Conditional Access coverage dashboard to the configured output directory.",
  fields: [
    {
      id: "includeDisabledPolicies",
      label: "Include disabled policies",
      type: "checkbox",
      defaultValue: false
    },
    tenantField
  ]
};

const legacyAuthExposureReportScript = {
  id: "m365-legacy-auth-exposure-report",
  name: "M365 Legacy Auth Exposure Report",
  category: "Identity",
  summary: "Review recent sign-ins tied to legacy authentication clients and protocols.",
  description:
    "Exports an HTML dashboard with recent legacy-auth sign-ins, affected users, and Conditional Access status visibility.",
  scriptRelativePath: "M365-LegacyAuthExposureReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-legacy-auth-exposure-report",
  outputs: "Writes an HTML legacy-auth exposure dashboard to the configured output directory.",
  fields: [
    {
      id: "lookbackDays",
      label: "Lookback window (days)",
      type: "number",
      required: false,
      defaultValue: 30,
      min: 7,
      max: 180
    },
    tenantField
  ]
};

const pimRoleActivationReportScript = {
  id: "m365-pim-role-activation-report",
  name: "M365 PIM Role Activation Report",
  category: "Identity",
  summary: "Review active and eligible privileged role schedule instances from Entra ID.",
  description:
    "Exports an HTML dashboard with active and eligible role schedules, including permanent assignments that may need review.",
  scriptRelativePath: "M365-PIMRoleActivationReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-pim-role-activation-report",
  outputs: "Writes an HTML PIM role activation dashboard to the configured output directory.",
  fields: [tenantField]
};

const deviceComplianceSnapshotScript = {
  id: "m365-device-compliance-snapshot",
  name: "M365 Device Compliance Snapshot",
  category: "Devices",
  summary: "Review managed device compliance state, ownership, and platform mix.",
  description:
    "Exports an HTML dashboard summarizing Intune-managed device compliance and ownership posture from Microsoft Graph.",
  scriptRelativePath: "M365-DeviceComplianceSnapshot.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-device-compliance-snapshot",
  outputs: "Writes an HTML device-compliance dashboard to the configured output directory.",
  fields: [tenantField]
};

const b2bDirectConnectReportScript = {
  id: "m365-b2b-direct-connect-report",
  name: "M365 B2B Direct Connect Report",
  category: "Identity",
  summary: "Review cross-tenant access defaults and B2B direct connect posture.",
  description:
    "Exports an HTML dashboard summarizing cross-tenant access policy defaults and partner posture from Microsoft Entra ID.",
  scriptRelativePath: "M365-B2BDirectConnectReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-b2b-direct-connect-report",
  outputs: "Writes an HTML B2B direct connect dashboard to the configured output directory.",
  fields: [tenantField]
};

const mailTransportRulesAuditScript = {
  id: "m365-mail-transport-rules-audit",
  name: "M365 Mail Transport Rules Audit",
  category: "Exchange",
  summary: "Review Exchange Online transport rules, modes, and mail-flow logic.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard for transport rule state, priority, and test/audit mode coverage.",
  scriptRelativePath: "M365-MailTransportRulesAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-mail-transport-rules-audit",
  outputs: "Writes an HTML transport-rules dashboard to the configured output directory.",
  fields: [tenantField]
};

const defenderIncidentSnapshotScript = {
  id: "m365-defender-incident-snapshot",
  name: "M365 Defender Incident Snapshot",
  category: "Security",
  summary: "Capture current incidents from Microsoft Defender via Graph Security.",
  description:
    "Exports an HTML dashboard with active Defender incidents, severity, status, and classification.",
  scriptRelativePath: "M365-DefenderIncidentSnapshot.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-defender-incident-snapshot",
  outputs: "Writes an HTML Defender incident dashboard to the configured output directory.",
  fields: [tenantField]
};

const privilegedGroupAuditScript = {
  id: "m365-privileged-group-audit",
  name: "M365 Privileged Group Audit",
  category: "Identity",
  summary: "Review sensitive groups for owner gaps, membership exposure, and visibility posture.",
  description:
    "Exports an HTML dashboard covering selected privileged groups, member counts, owner state, and visibility.",
  scriptRelativePath: "M365-PrivilegedGroupAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-privileged-group-audit",
  outputs: "Writes an HTML privileged-group dashboard to the configured output directory.",
  fields: [tenantField]
};

const passwordResetReadinessReportScript = {
  id: "m365-password-reset-readiness-report",
  name: "M365 Password Reset Readiness Report",
  category: "Identity",
  summary: "Review SSPR readiness based on registered password-reset-capable authentication methods.",
  description:
    "Exports an HTML dashboard that highlights which users appear ready for self-service password reset and which still need method registration.",
  scriptRelativePath: "M365-PasswordResetReadinessReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-password-reset-readiness-report",
  outputs: "Writes an HTML password-reset readiness dashboard to the configured output directory.",
  fields: [tenantField]
};

const oneDriveExternalSharingReportScript = {
  id: "m365-onedrive-external-sharing-report",
  name: "M365 OneDrive External Sharing Report",
  category: "SharePoint",
  summary: "Review OneDrive site usage and surface large or highly active personal sites.",
  description:
    "Exports an HTML dashboard with OneDrive usage data, site sizes, and recent activity to support sharing and storage reviews.",
  scriptRelativePath: "M365-OneDriveExternalSharingReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-onedrive-external-sharing-report",
  outputs: "Writes an HTML OneDrive sharing dashboard to the configured output directory.",
  fields: [tenantField]
};

export const scripts = [
  compromisedAccountScript,
  checkMfaStatusScript,
  usageReportScript,
  licensingReportScript,
  guestAccessReportScript,
  conditionalAccessReportScript,
  mailForwardingAuditScript,
  sharedMailboxReportScript,
  signInRiskReportScript,
  teamsExternalAccessReportScript,
  sharePointSharingReportScript,
  secureScoreSnapshotScript,
  adminRoleAuditScript,
  inactiveUsersReportScript,
  appConsentAuditScript,
  mailboxPermissionAuditScript,
  externalSharingLinksReportScript,
  distributionGroupAuditScript,
  serviceHealthSnapshotScript,
  authenticationPolicyReportScript,
  privilegedAppAuditScript,
  dkimDmarcReportScript,
  groupLifecycleReportScript,
  caPolicyCoverageReportScript,
  legacyAuthExposureReportScript,
  pimRoleActivationReportScript,
  deviceComplianceSnapshotScript,
  b2bDirectConnectReportScript,
  mailTransportRulesAuditScript,
  defenderIncidentSnapshotScript,
  privilegedGroupAuditScript,
  passwordResetReadinessReportScript,
  oneDriveExternalSharingReportScript
];
