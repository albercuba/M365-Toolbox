const compromisedAccountScript = {
  id: "m365-compromised-account-remediation",
  name: "Compromised Account Remediation",
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
    {
      id: "tenantId",
      label: "Tenant ID or Domain",
      type: "text",
      required: false,
      placeholder: "contoso.onmicrosoft.com",
      helpText: "Optional. If provided, device-code sign-in is scoped to this tenant."
    },
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
  scriptRelativePath: "Get-M365MfaReport.ps1",
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
    {
      id: "tenantId",
      label: "Tenant ID or Domain",
      type: "text",
      required: false,
      placeholder: "contoso.onmicrosoft.com",
      helpText: "Recommended when Microsoft Graph cannot infer the tenant automatically."
    }
  ]
};

const usageReportScript = {
  id: "m365-usage-report",
  name: "M365 Usage Report",
  category: "Reporting",
  summary: "Generate OneDrive, SharePoint, and Mailbox storage usage dashboards.",
  description:
    "Runs the approved Microsoft 365 storage usage workflow with Microsoft Graph device-code authentication and exports an HTML dashboard into the toolbox output folder.",
  scriptRelativePath: "Get-M365UsageReport.ps1",
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
    {
      id: "tenantId",
      label: "Tenant ID or Domain",
      type: "text",
      required: false,
      placeholder: "contoso.onmicrosoft.com",
      helpText: "Optional. If provided, device-code sign-in is scoped to this tenant."
    }
  ]
};

export const scripts = [compromisedAccountScript, checkMfaStatusScript, usageReportScript];
