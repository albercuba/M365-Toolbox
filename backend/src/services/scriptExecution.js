import path from "node:path";
import fs from "node:fs";
import { OUTPUT_DIR, TOOLBOX_SCRIPT_MOUNT_ROOT } from "../config/runtime.js";
import { createError, normalizeListValue } from "./validation.js";

function getScriptMountRoot(script) {
  if (script.scriptMountRootEnv && process.env[script.scriptMountRootEnv]) {
    return process.env[script.scriptMountRootEnv];
  }

  return TOOLBOX_SCRIPT_MOUNT_ROOT;
}

function findScriptByFileName(rootPath, fileName) {
  if (!fs.existsSync(rootPath)) {
    return null;
  }

  const entries = fs.readdirSync(rootPath, { withFileTypes: true });
  for (const entry of entries) {
    const entryPath = path.posix.join(rootPath.replace(/\\/g, "/"), entry.name);
    if (entry.isFile() && entry.name === fileName) {
      return entryPath;
    }
    if (entry.isDirectory()) {
      const nestedMatch = findScriptByFileName(entryPath, fileName);
      if (nestedMatch) {
        return nestedMatch;
      }
    }
  }

  return null;
}

export function resolveScriptPath(script) {
  const scriptMountRoot = getScriptMountRoot(script).replace(/\\/g, "/");
  const configuredPath = path.posix.join(scriptMountRoot, script.scriptRelativePath.replace(/\\/g, "/"));

  if (fs.existsSync(configuredPath)) {
    return configuredPath;
  }

  const fallbackPath = findScriptByFileName(scriptMountRoot, path.posix.basename(script.scriptRelativePath));
  if (fallbackPath) {
    return fallbackPath;
  }

  throw createError(`Script file not found in container. Expected '${configuredPath}'.`, 500);
}

function buildCompromisedAccountArgs(script, payload) {
  const timestamp = new Date().toISOString().replace(/[:]/g, "-");
  const artifactBase = path.posix.join(OUTPUT_DIR.replace(/\\/g, "/"), `m365-compromised-account-remediation-${timestamp}`);
  const args = ["-OutputPath", OUTPUT_DIR, "-ExportHtml", `${artifactBase}.html`];
  const upns = normalizeListValue(payload.userPrincipalName);
  const actions = normalizeListValue(payload.actions);

  if (upns.length > 0) args.push("-UserPrincipalName", upns.join(","));
  if (actions.length > 0) args.push("-Actions", actions.join(","));
  if (payload.auditLogDays) args.push("-AuditLogDays", String(payload.auditLogDays));
  if (payload.tenantId) args.push("-TenantId", String(payload.tenantId));
  if (payload.includeGeneratedPasswordsInResults) args.push("-IncludeGeneratedPasswordsInResults");
  if (payload.exportIncidentPackage) args.push("-ExportIncidentPackage");
  if (payload.whatIf) args.push("-WhatIf");

  return {
    args,
    artifacts: {
      basePath: artifactBase,
      htmlPath: `${artifactBase}.html`,
      files: []
    }
  };
}

function buildMfaStatusArgs(script, payload) {
  const timestamp = new Date().toISOString().replace(/[:]/g, "-");
  const artifactBase = path.posix.join(OUTPUT_DIR.replace(/\\/g, "/"), `m365-mfa-report-${timestamp}`);
  const args = [];

  if (payload.includeGuests) args.push("-IncludeGuests");
  if (payload.tenantId) args.push("-TenantId", String(payload.tenantId));
  if (payload.exportHtml !== false) args.push("-ExportHtml", `${artifactBase}.html`);

  return {
    args,
    artifacts: {
      basePath: artifactBase,
      htmlPath: payload.exportHtml !== false ? `${artifactBase}.html` : null,
      files: []
    }
  };
}

function buildUsageReportArgs(script, payload) {
  const timestamp = new Date().toISOString().replace(/[:]/g, "-");
  const artifactBase = path.posix.join(OUTPUT_DIR.replace(/\\/g, "/"), `m365-usage-report-${timestamp}`);
  const args = ["-OutputPath", OUTPUT_DIR, "-ExportHtml", `${artifactBase}.html`];
  const reports = normalizeListValue(payload.reports);

  if (reports.length > 0) args.push("-Reports", reports.join(","));
  if (payload.tenantId) args.push("-TenantId", String(payload.tenantId));

  return {
    args,
    artifacts: {
      basePath: artifactBase,
      htmlPath: `${artifactBase}.html`,
      files: []
    }
  };
}

function buildGenericHtmlArgs(script, payload) {
  const timestamp = new Date().toISOString().replace(/[:]/g, "-");
  const artifactBase = path.posix.join(
    OUTPUT_DIR.replace(/\\/g, "/"),
    `${script.outputBaseName || script.id}-${timestamp}`
  );
  const args = ["-OutputPath", OUTPUT_DIR, "-ExportHtml", `${artifactBase}.html`];

  for (const field of script.fields || []) {
    const rawValue = payload[field.id];
    const paramName = field.paramName || field.id;

    if (field.type === "checkbox") {
      if (rawValue) {
        args.push(`-${paramName}`);
      }
      continue;
    }

    if (field.type === "multiselect") {
      const items = normalizeListValue(rawValue);
      if (items.length > 0) {
        args.push(`-${paramName}`, items.join(","));
      }
      continue;
    }

    if (rawValue === undefined || rawValue === null || rawValue === "") {
      continue;
    }

    args.push(`-${paramName}`, String(rawValue));
  }

  return {
    args,
    artifacts: {
      basePath: artifactBase,
      htmlPath: `${artifactBase}.html`,
      files: []
    }
  };
}

export function buildExecutionPlan(script, payload) {
  const scriptPath = resolveScriptPath(script);
  const wrapperPath = path.posix.join(TOOLBOX_SCRIPT_MOUNT_ROOT.replace(/\\/g, "/"), "Invoke-ToolboxScript.ps1");
  const runtimeScript = { ...script, scriptPath };
  let result;

  switch (script.id) {
    case "m365-compromised-account-remediation":
      result = buildCompromisedAccountArgs(runtimeScript, payload);
      break;
    case "m365-check-mfa-status":
      result = buildMfaStatusArgs(runtimeScript, payload);
      break;
    case "m365-usage-report":
      result = buildUsageReportArgs(runtimeScript, payload);
      break;
    default:
      if (script.runner === "generic-html") {
        result = buildGenericHtmlArgs(runtimeScript, payload);
        break;
      }
      throw createError(`No runner is defined for script '${script.id}'.`, 500);
  }

  return {
    scriptPath,
    artifacts: result.artifacts,
    commandArgs: [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      wrapperPath,
      "-ScriptPath",
      scriptPath,
      ...result.args
    ]
  };
}
