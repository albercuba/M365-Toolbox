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
  name: "Check MFA Status",
  category: "Identity",
  summary: "Review MFA registration coverage and registered authentication methods.",
  description:
    "Runs the approved MFA status report across the tenant and can export the results as a CSV into the toolbox output folder.",
  scriptRelativePath: "M365 - CompromisedAccountRemediation/Check-MFAStatus.ps1",
  outputs:
    "Writes an MFA status CSV export to the configured output directory when export is enabled.",
  fields: [
    {
      id: "includeGuests",
      label: "Include guest users",
      type: "checkbox",
      defaultValue: false
    },
    {
      id: "skipDisabled",
      label: "Skip disabled users",
      type: "checkbox",
      defaultValue: true
    },
    {
      id: "exportCsv",
      label: "Export CSV",
      type: "checkbox",
      defaultValue: true
    }
  ]
};

export const scripts = [compromisedAccountScript, checkMfaStatusScript];
