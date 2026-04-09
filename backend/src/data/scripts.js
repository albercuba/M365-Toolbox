const tenantField = {
  id: "tenantId",
  label: "Tenant ID or Domain",
  type: "text",
  required: false,
  placeholder: "contoso.onmicrosoft.com",
  helpText: "Optional. If provided, device-code sign-in is scoped to this tenant."
};

const destructiveScriptIds = new Set([
  "m365-compromised-account-remediation"
]);

function withCommonMetadata(script) {
  const mode = destructiveScriptIds.has(script.id) ? "remediation" : "read-only";
  const estimatedRuntimeMinutes = script.id === "m365-compromised-account-remediation"
    ? 8
    : script.category === "Exchange" || script.category === "SharePoint" || script.category === "Teams"
      ? 4
      : 3;

  return {
    ...script,
    mode,
    approvalRequired: mode === "remediation",
    prerequisites: [
      "PowerShell 7 and backend Microsoft Graph modules available in the execution container",
      "Microsoft 365 admin sign-in by device code when prompted",
      script.category === "Exchange" ? "Exchange Online permissions for the chosen account" : "Microsoft Graph permissions required by the selected workflow"
    ],
    permissions: mode === "remediation"
      ? ["Read tenant data", "Modify tenant state", "Generate report artifacts"]
      : ["Read tenant data", "Generate report artifacts"],
    estimatedRuntimeMinutes,
    examples: [
      {
        title: "Default tenant-wide run",
        description: "Use the default field values and authenticate when the device-code prompt appears."
      },
      {
        title: "Scoped admin review",
        description: "Provide a tenant id or domain to scope authentication to the expected Microsoft 365 tenant."
      }
    ]
  };
}

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
  reviewActions: [
    "ReviewMfaMethods",
    "ReviewInboxRules",
    "ReviewMailboxForwarding",
    "ReviewMailboxDelegates",
    "ReviewRecentSignIns",
    "ExportAuditLog"
  ],
  highImpactActions: [
    "DisableUser",
    "ResetPassword",
    "RemoveMfaMethods",
    "RemoveMailboxForwarding",
    "RemoveMailboxDelegates",
    "DisableMailboxProtocols"
  ],
  actionProfiles: [
    {
      id: "review-only",
      label: "Review Only",
      description: "Collect recent sign-ins, MFA, inbox rule, forwarding, delegate, and audit evidence without changing tenant state.",
      actions: [
        "ReviewMfaMethods",
        "ReviewInboxRules",
        "ReviewMailboxForwarding",
        "ReviewMailboxDelegates",
        "ReviewRecentSignIns",
        "ExportAuditLog"
      ]
    },
    {
      id: "containment",
      label: "Containment",
      description: "Apply a fast containment bundle for active account compromise, including sign-out, password reset, mailbox cleanup, and protocol lock-down.",
      actions: [
        "DisableUser",
        "RevokeSessions",
        "ResetPassword",
        "RemoveMfaMethods",
        "DisableInboxRules",
        "RemoveMailboxForwarding",
        "DisableSignature",
        "DisableMailboxProtocols",
        "ExportAuditLog"
      ]
    }
  ],
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
        "ReviewInboxRules",
        "DisableInboxRules",
        "ReviewMailboxForwarding",
        "RemoveMailboxForwarding",
        "DisableSignature",
        "ReviewMailboxDelegates",
        "RemoveMailboxDelegates",
        "ReviewRecentSignIns",
        "DisableMailboxProtocols",
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
      id: "exportIncidentPackage",
      label: "Export per-user incident package",
      type: "checkbox",
      defaultValue: true,
      helpText: "Creates a predictable folder per target with supporting CSV, log, and summary files."
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

const mfaRegistrationCampaignReportScript = {
  id: "m365-mfa-registration-campaign-report",
  name: "M365 MFA Registration Campaign Report",
  category: "Identity",
  summary: "Track MFA registration readiness by user and department for adoption campaigns.",
  description:
    "Exports an HTML dashboard that highlights who is and is not registered for MFA, along with department-level campaign targeting visibility.",
  scriptRelativePath: "M365-MFARegistrationCampaignReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-mfa-registration-campaign-report",
  outputs: "Writes an HTML MFA registration campaign dashboard to the configured output directory.",
  fields: [
    {
      id: "maxUsersToInspect",
      label: "Maximum users to inspect",
      type: "number",
      required: false,
      defaultValue: 250,
      min: 50,
      max: 1000
    },
    tenantField
  ]
};

const enterpriseAppsInventoryScript = {
  id: "m365-enterprise-apps-inventory",
  name: "M365 Enterprise Apps Inventory",
  category: "Security",
  summary: "Inventory enterprise apps with publisher, assignment, and credential visibility.",
  description:
    "Exports an HTML dashboard that inventories service principals, assignment requirements, publishers, and credential counts.",
  scriptRelativePath: "M365-EnterpriseAppsInventory.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-enterprise-apps-inventory",
  outputs: "Writes an HTML enterprise-app inventory dashboard to the configured output directory.",
  fields: [tenantField]
};

const guestInvitationFailuresScript = {
  id: "m365-guest-invitation-failures",
  name: "M365 Guest Invitation Failures",
  category: "Identity",
  summary: "Find guest invitations that were never accepted or never resulted in sign-in activity.",
  description:
    "Exports an HTML dashboard of guest accounts with pending invitation state, stale acceptance, or no successful sign-in history.",
  scriptRelativePath: "M365-GuestInvitationFailures.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-guest-invitation-failures",
  outputs: "Writes an HTML guest invitation dashboard to the configured output directory.",
  fields: [
    {
      id: "staleDays",
      label: "Stale threshold (days)",
      type: "number",
      required: false,
      defaultValue: 30,
      min: 7,
      max: 365
    },
    tenantField
  ]
};

const mailboxAutoReplyAuditScript = {
  id: "m365-mailbox-auto-reply-audit",
  name: "M365 Mailbox Auto-Reply Audit",
  category: "Exchange",
  summary: "Review automatic reply configuration across user and shared mailboxes.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard for active and external automatic replies.",
  scriptRelativePath: "M365-MailboxAutoReplyAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-mailbox-auto-reply-audit",
  outputs: "Writes an HTML mailbox auto-reply dashboard to the configured output directory.",
  fields: [tenantField]
};

const calendarSharingAuditScript = {
  id: "m365-calendar-sharing-audit",
  name: "M365 Calendar Sharing Audit",
  category: "Exchange",
  summary: "Review calendar delegate and anonymous sharing exposure across inspected mailboxes.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard for calendar sharing posture and anonymous exposure.",
  scriptRelativePath: "M365-CalendarSharingAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-calendar-sharing-audit",
  outputs: "Writes an HTML calendar sharing dashboard to the configured output directory.",
  fields: [
    {
      id: "maxMailboxesToInspect",
      label: "Maximum mailboxes to inspect",
      type: "number",
      required: false,
      defaultValue: 150,
      min: 25,
      max: 500
    },
    tenantField
  ]
};

const roleEligibleAssignmentsReportScript = {
  id: "m365-role-eligible-assignments-report",
  name: "M365 Role Eligible Assignments Report",
  category: "Identity",
  summary: "Review eligible privileged roles, principals, and permanent assignment patterns.",
  description:
    "Exports an HTML dashboard for Entra ID role eligibility schedule instances, including assignment timing and permanence.",
  scriptRelativePath: "M365-RoleEligibleAssignmentsReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-role-eligible-assignments-report",
  outputs: "Writes an HTML eligible-role dashboard to the configured output directory.",
  fields: [tenantField]
};

const anonymousLinkExposureReportScript = {
  id: "m365-anonymous-link-exposure-report",
  name: "M365 Anonymous Link Exposure Report",
  category: "SharePoint",
  summary: "Review SharePoint sharing defaults and site footprint that can amplify anyone-link exposure.",
  description:
    "Exports an HTML dashboard that combines tenant sharing settings with SharePoint site usage inventory.",
  scriptRelativePath: "M365-AnonymousLinkExposureReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-anonymous-link-exposure-report",
  outputs: "Writes an HTML anonymous-link exposure dashboard to the configured output directory.",
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

const teamsOwnershipAuditScript = {
  id: "m365-teams-ownership-audit",
  name: "M365 Teams Ownership Audit",
  category: "Teams",
  summary: "Find ownerless or single-owner Teams and highlight guest-heavy ownership risk.",
  description:
    "Exports an HTML dashboard covering Teams ownership resilience, owner names, and guest concentration.",
  scriptRelativePath: "M365-TeamsOwnershipAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-teams-ownership-audit",
  outputs: "Writes an HTML Teams ownership dashboard to the configured output directory.",
  fields: [
    {
      id: "maxTeamsToInspect",
      label: "Maximum teams to inspect",
      type: "number",
      required: false,
      defaultValue: 150,
      min: 25,
      max: 500
    },
    tenantField
  ]
};

const appCredentialExpiryReportScript = {
  id: "m365-app-credential-expiry-report",
  name: "M365 App Credential Expiry Report",
  category: "Security",
  summary: "Review expiring and expired secrets or certificates across app registrations.",
  description:
    "Exports an HTML dashboard that highlights app credentials nearing expiry or already expired.",
  scriptRelativePath: "M365-AppCredentialExpiryReport.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-app-credential-expiry-report",
  outputs: "Writes an HTML app credential expiry dashboard to the configured output directory.",
  fields: [
    {
      id: "daysAhead",
      label: "Expiry window (days)",
      type: "number",
      required: false,
      defaultValue: 60,
      min: 7,
      max: 365
    },
    tenantField
  ]
};

const mailflowConnectorAuditScript = {
  id: "m365-mailflow-connector-audit",
  name: "M365 Mailflow Connector Audit",
  category: "Exchange",
  summary: "Review inbound and outbound Exchange Online connectors, relay paths, and TLS posture.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard for mailflow connector inventory and status.",
  scriptRelativePath: "M365-MailflowConnectorAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-mailflow-connector-audit",
  outputs: "Writes an HTML mailflow connector dashboard to the configured output directory.",
  fields: [tenantField]
};

const breakGlassAccountAuditScript = {
  id: "m365-break-glass-account-audit",
  name: "M365 Break-Glass Account Audit",
  category: "Identity",
  summary: "Review emergency access accounts for privilege, MFA, and recent sign-in visibility.",
  description:
    "Exports an HTML dashboard that locates likely break-glass accounts and checks whether they are enabled, privileged, and protected by MFA.",
  scriptRelativePath: "M365-BreakGlassAccountAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-break-glass-account-audit",
  outputs: "Writes an HTML break-glass account dashboard to the configured output directory.",
  fields: [tenantField]
};

const impossibleTravelReviewScript = {
  id: "m365-impossible-travel-review",
  name: "M365 Impossible Travel Review",
  category: "Identity",
  summary: "Review sign-ins that jump between countries too quickly to be expected.",
  description:
    "Exports an HTML dashboard highlighting rapid country changes in recent Microsoft Entra sign-ins within the selected time threshold.",
  scriptRelativePath: "M365-ImpossibleTravelReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-impossible-travel-review",
  outputs: "Writes an HTML impossible-travel dashboard to the configured output directory.",
  fields: [
    {
      id: "lookbackDays",
      label: "Lookback window (days)",
      type: "number",
      required: false,
      defaultValue: 14,
      min: 1,
      max: 90
    },
    {
      id: "maxHoursBetween",
      label: "Maximum hours between sign-ins",
      type: "number",
      required: false,
      defaultValue: 12,
      min: 1,
      max: 48
    },
    tenantField
  ]
};

const mailboxLoginAnomalyReviewScript = {
  id: "m365-mailbox-login-anomaly-review",
  name: "M365 Mailbox Login Anomaly Review",
  category: "Exchange",
  summary: "Review mailbox-related sign-ins that use suspicious clients, fail Conditional Access, or look unusual.",
  description:
    "Exports an HTML dashboard of Exchange and Outlook sign-ins with anomaly signals such as legacy client usage, failures, and non-interactive mailbox access.",
  scriptRelativePath: "M365-MailboxLoginAnomalyReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-mailbox-login-anomaly-review",
  outputs: "Writes an HTML mailbox login anomaly dashboard to the configured output directory.",
  fields: [
    {
      id: "lookbackDays",
      label: "Lookback window (days)",
      type: "number",
      required: false,
      defaultValue: 14,
      min: 1,
      max: 90
    },
    tenantField
  ]
};

const externalTenantTrustReviewScript = {
  id: "m365-external-tenant-trust-review",
  name: "M365 External Tenant Trust Review",
  category: "Identity",
  summary: "Review cross-tenant access partners and external guest domain concentration.",
  description:
    "Exports an HTML dashboard summarizing cross-tenant partner trust settings and the external domains most represented in guest access.",
  scriptRelativePath: "M365-ExternalTenantTrustReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-external-tenant-trust-review",
  outputs: "Writes an HTML external tenant trust dashboard to the configured output directory.",
  fields: [tenantField]
};

const privilegedUserSignInReviewScript = {
  id: "m365-privileged-user-sign-in-review",
  name: "M365 Privileged User Sign-In Review",
  category: "Identity",
  summary: "Review recent sign-in visibility for privileged accounts across active directory roles.",
  description:
    "Exports an HTML dashboard listing privileged users, their roles, and their most recent sign-in activity.",
  scriptRelativePath: "M365-PrivilegedUserSignInReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-privileged-user-sign-in-review",
  outputs: "Writes an HTML privileged user sign-in dashboard to the configured output directory.",
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

const inboxRuleThreatHuntScript = {
  id: "m365-inbox-rule-threat-hunt",
  name: "M365 Inbox Rule Threat Hunt",
  category: "Exchange",
  summary: "Hunt for suspicious inbox rules that forward, redirect, hide, or delete messages.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard for inbox rules with threat-hunting signals.",
  scriptRelativePath: "M365-InboxRuleThreatHunt.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-inbox-rule-threat-hunt",
  outputs: "Writes an HTML inbox rule threat-hunt dashboard to the configured output directory.",
  fields: [
    {
      id: "maxMailboxesToInspect",
      label: "Maximum mailboxes to inspect",
      type: "number",
      required: false,
      defaultValue: 150,
      min: 25,
      max: 500
    },
    tenantField
  ]
};

const conditionalAccessGapReviewScript = {
  id: "m365-conditional-access-gap-review",
  name: "M365 Conditional Access Gap Review",
  category: "Identity",
  summary: "Review Conditional Access policies for exclusions, missing grant controls, and non-enforced states.",
  description:
    "Exports an HTML dashboard highlighting Conditional Access policies that may have coverage or enforcement gaps.",
  scriptRelativePath: "M365-ConditionalAccessGapReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-conditional-access-gap-review",
  outputs: "Writes an HTML Conditional Access gap dashboard to the configured output directory.",
  fields: [tenantField]
};

const mfaExclusionAuditScript = {
  id: "m365-mfa-exclusion-audit",
  name: "M365 MFA Exclusion Audit",
  category: "Identity",
  summary: "Audit MFA-enforcing policies that still exclude users, groups, or roles.",
  description:
    "Exports an HTML dashboard focused on MFA-related Conditional Access policies and their exclusion footprint.",
  scriptRelativePath: "M365-MFAExclusionAudit.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-mfa-exclusion-audit",
  outputs: "Writes an HTML MFA exclusion dashboard to the configured output directory.",
  fields: [tenantField]
};

const dormantAdminAccountReviewScript = {
  id: "m365-dormant-admin-account-review",
  name: "M365 Dormant Admin Account Review",
  category: "Identity",
  summary: "Find privileged accounts with little or no recent sign-in activity.",
  description:
    "Exports an HTML dashboard that highlights dormant privileged accounts based on configurable sign-in thresholds.",
  scriptRelativePath: "M365-DormantAdminAccountReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-dormant-admin-account-review",
  outputs: "Writes an HTML dormant admin account dashboard to the configured output directory.",
  fields: [
    {
      id: "dormantDays",
      label: "Dormant threshold (days)",
      type: "number",
      required: false,
      defaultValue: 45,
      min: 7,
      max: 365
    },
    {
      id: "lookbackDays",
      label: "Lookback window (days)",
      type: "number",
      required: false,
      defaultValue: 90,
      min: 14,
      max: 365
    },
    tenantField
  ]
};

const sharedMailboxAbuseReviewScript = {
  id: "m365-shared-mailbox-abuse-review",
  name: "M365 Shared Mailbox Abuse Review",
  category: "Exchange",
  summary: "Review shared mailboxes for forwarding, risky delegates, and inbox rule abuse signals.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard for shared mailbox abuse indicators.",
  scriptRelativePath: "M365-SharedMailboxAbuseReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-shared-mailbox-abuse-review",
  outputs: "Writes an HTML shared mailbox abuse dashboard to the configured output directory.",
  fields: [
    {
      id: "maxMailboxesToInspect",
      label: "Maximum mailboxes to inspect",
      type: "number",
      required: false,
      defaultValue: 150,
      min: 25,
      max: 500
    },
    tenantField
  ]
};

const transportRuleThreatHuntScript = {
  id: "m365-transport-rule-threat-hunt",
  name: "M365 Transport Rule Threat Hunt",
  category: "Exchange",
  summary: "Hunt for transport rules that redirect, bypass filtering, or manipulate mail flow in risky ways.",
  description:
    "Connects to Exchange Online by device code and exports an HTML dashboard for transport rules with threat-hunting indicators.",
  scriptRelativePath: "M365-TransportRuleThreatHunt.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-transport-rule-threat-hunt",
  outputs: "Writes an HTML transport rule threat-hunt dashboard to the configured output directory.",
  fields: [tenantField]
};

const breakGlassAccountHardeningReviewScript = {
  id: "m365-break-glass-account-hardening-review",
  name: "M365 Break-Glass Account Hardening Review",
  category: "Identity",
  summary: "Review likely emergency access accounts for MFA, state, and recent sign-in posture.",
  description:
    "Exports an HTML dashboard that searches for likely break-glass accounts and summarizes their hardening state.",
  scriptRelativePath: "M365-BreakGlassAccountHardeningReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-break-glass-account-hardening-review",
  outputs: "Writes an HTML break-glass hardening dashboard to the configured output directory.",
  fields: [
    {
      id: "accountNameHints",
      label: "Account name hints",
      type: "text",
      required: false,
      defaultValue: "breakglass,emergency",
      placeholder: "breakglass, emergency, backupadmin"
    },
    tenantField
  ]
};

const oauthAppRiskReviewScript = {
  id: "m365-oauth-app-risk-review",
  name: "M365 OAuth App Risk Review",
  category: "Security",
  summary: "Review OAuth-enabled apps for delegated grants, elevated scopes, and credential risk.",
  description:
    "Exports an HTML dashboard highlighting enterprise apps with delegated consent, risky scopes, and expiring credentials.",
  scriptRelativePath: "M365-OAuthAppRiskReview.ps1",
  scriptMountRootEnv: "TOOLBOX_SCRIPT_MOUNT_ROOT",
  runner: "generic-html",
  outputBaseName: "m365-oauth-app-risk-review",
  outputs: "Writes an HTML OAuth app risk dashboard to the configured output directory.",
  fields: [
    {
      id: "daysAhead",
      label: "Credential expiry window (days)",
      type: "number",
      required: false,
      defaultValue: 60,
      min: 7,
      max: 365
    },
    tenantField
  ]
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
  oneDriveExternalSharingReportScript,
  mfaRegistrationCampaignReportScript,
  enterpriseAppsInventoryScript,
  guestInvitationFailuresScript,
  mailboxAutoReplyAuditScript,
  calendarSharingAuditScript,
  roleEligibleAssignmentsReportScript,
  anonymousLinkExposureReportScript,
  teamsOwnershipAuditScript,
  appCredentialExpiryReportScript,
  mailflowConnectorAuditScript,
  breakGlassAccountAuditScript,
  impossibleTravelReviewScript,
  mailboxLoginAnomalyReviewScript,
  externalTenantTrustReviewScript,
  privilegedUserSignInReviewScript,
  inboxRuleThreatHuntScript,
  conditionalAccessGapReviewScript,
  mfaExclusionAuditScript,
  dormantAdminAccountReviewScript,
  sharedMailboxAbuseReviewScript,
  transportRuleThreatHuntScript,
  breakGlassAccountHardeningReviewScript,
  oauthAppRiskReviewScript
].map(withCommonMetadata);
