const compromisedAccountScript = {
  id: "m365-compromised-account-remediation",
  name: "Compromised Account Remediation",
  category: "Incident Response",
  summary: "Contain a compromised Microsoft 365 account and export investigation data.",
  description:
    "Runs the approved M365 compromised-account workflow against one or more target users. The runner only exposes allowlisted parameters so the web UI can stay safe and predictable.",
  scriptRelativePath:
    "M365 - CompromisedAccountRemediation/M365-CompromisedAccountRemediation.ps1",
  outputs:
    "Writes status CSVs, error logs, unified audit exports, sign-in exports, and message trace reports to the configured output directory.",
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
      id: "csvPath",
      label: "CSV Path",
      type: "text",
      required: false,
      placeholder: "/workspace-scripts/path/to/users.csv",
      helpText: "Optional path inside the mounted scripts folder."
    },
    {
      id: "tenantId",
      label: "Tenant ID",
      type: "text",
      required: false
    },
    {
      id: "clientId",
      label: "Client ID",
      type: "text",
      required: false
    },
    {
      id: "certificateThumbprint",
      label: "Certificate Thumbprint",
      type: "text",
      required: false
    },
    {
      id: "installMissingModules",
      label: "Install missing modules",
      type: "checkbox",
      defaultValue: false
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
    },
    {
      id: "exportHtml",
      label: "Export HTML dashboard",
      type: "checkbox",
      defaultValue: true
    },
    {
      id: "exportXlsx",
      label: "Export Excel report",
      type: "checkbox",
      defaultValue: true
    }
  ]
};

export const scripts = [compromisedAccountScript, checkMfaStatusScript];
