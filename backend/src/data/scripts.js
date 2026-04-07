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
  adminRoleAuditScript
];
