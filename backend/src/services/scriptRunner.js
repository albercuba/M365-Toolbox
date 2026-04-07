import { spawn } from "node:child_process";
import path from "node:path";
import fs from "node:fs";
import { v4 as uuidv4 } from "uuid";
import { scripts } from "../data/scripts.js";

const OUTPUT_DIR = process.env.OUTPUT_DIR || path.resolve(process.cwd(), "../output");
const SCRIPT_MOUNT_ROOT = process.env.SCRIPT_MOUNT_ROOT || "C:/VSCode/Powershell";
const TOOLBOX_SCRIPT_MOUNT_ROOT = process.env.TOOLBOX_SCRIPT_MOUNT_ROOT || path.resolve(process.cwd(), "../scripts");
const runStore = new Map();

function updateArtifactsFromStdout(run) {
  if (!run?.stdout) {
    return;
  }

  const htmlMatch = run.stdout.match(/\[\+\]\s+HTML dashboard exported to:\s+(.+)/i);
  if (htmlMatch?.[1]) {
    run.artifacts = {
      ...run.artifacts,
      htmlPath: htmlMatch[1].trim()
    };
  }
}

function ensureOutputDir() {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

function normalizeListValue(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.map((item) => String(item).trim()).filter(Boolean);
  return String(value)
    .split(/[\n,]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function resolveScript(scriptId) {
  const script = scripts.find((entry) => entry.id === scriptId);
  if (!script) {
    throw new Error(`Unknown script '${scriptId}'.`);
  }

  return script;
}

function getScriptMountRoot(script) {
  if (script.scriptMountRootEnv && process.env[script.scriptMountRootEnv]) {
    return process.env[script.scriptMountRootEnv];
  }

  return SCRIPT_MOUNT_ROOT;
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

function resolveScriptPath(script) {
  const scriptMountRoot = getScriptMountRoot(script).replace(/\\/g, "/");
  const configuredPath = path.posix.join(scriptMountRoot, script.scriptRelativePath.replace(/\\/g, "/"));

  if (fs.existsSync(configuredPath)) {
    return configuredPath;
  }

  const fallbackPath = findScriptByFileName(scriptMountRoot, path.posix.basename(script.scriptRelativePath));
  if (fallbackPath) {
    return fallbackPath;
  }

  throw new Error(`Script file not found in container. Expected '${configuredPath}'.`);
}

function buildCompromisedAccountArgs(script, payload) {
  const timestamp = new Date().toISOString().replace(/[:]/g, "-");
  const outputBase = path.posix.join(OUTPUT_DIR.replace(/\\/g, "/"), `m365-compromised-account-remediation-${timestamp}`);
  const args = ["-OutputPath", OUTPUT_DIR, "-ExportHtml", `${outputBase}.html`];

  const upns = normalizeListValue(payload.userPrincipalName);
  if (upns.length > 0) {
    args.push("-UserPrincipalName", upns.join(","));
  }

  const actions = normalizeListValue(payload.actions);
  if (actions.length > 0) {
    args.push("-Actions", actions.join(","));
  }

  if (payload.auditLogDays) {
    args.push("-AuditLogDays", String(payload.auditLogDays));
  }

  if (payload.tenantId) {
    args.push("-TenantId", String(payload.tenantId));
  }

  if (payload.includeGeneratedPasswordsInResults) {
    args.push("-IncludeGeneratedPasswordsInResults");
  }

  if (payload.whatIf) {
    args.push("-WhatIf");
  }

  return {
    args,
    artifacts: {
      htmlPath: `${outputBase}.html`
    }
  };
}

function buildMfaStatusArgs(script, payload) {
  const args = [];
  const timestamp = new Date().toISOString().replace(/[:]/g, "-");
  const outputBase = path.posix.join(OUTPUT_DIR.replace(/\\/g, "/"), `m365-mfa-report-${timestamp}`);

  if (payload.includeGuests) {
    args.push("-IncludeGuests");
  }

  if (payload.tenantId) {
    args.push("-TenantId", String(payload.tenantId));
  }

  if (payload.exportHtml !== false) {
    args.push("-ExportHtml", `${outputBase}.html`);
  }

  return {
    args,
    artifacts: {
      htmlPath: payload.exportHtml !== false ? `${outputBase}.html` : null,
      xlsxPath: null
    }
  };
}

function buildUsageReportArgs(script, payload) {
  const timestamp = new Date().toISOString().replace(/[:]/g, "-");
  const outputBase = path.posix.join(OUTPUT_DIR.replace(/\\/g, "/"), `m365-usage-report-${timestamp}`);
  const args = ["-OutputPath", OUTPUT_DIR, "-ExportHtml", `${outputBase}.html`];
  const reports = normalizeListValue(payload.reports);

  if (reports.length > 0) {
    args.push("-Reports", reports.join(","));
  }

  if (payload.tenantId) {
    args.push("-TenantId", String(payload.tenantId));
  }

  return {
    args,
    artifacts: {
      htmlPath: `${outputBase}.html`
    }
  };
}

function buildGenericHtmlArgs(script, payload) {
  const timestamp = new Date().toISOString().replace(/[:]/g, "-");
  const outputBase = path.posix.join(
    OUTPUT_DIR.replace(/\\/g, "/"),
    `${script.outputBaseName || script.id}-${timestamp}`
  );
  const args = ["-OutputPath", OUTPUT_DIR, "-ExportHtml", `${outputBase}.html`];

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
      htmlPath: `${outputBase}.html`
    }
  };
}

function buildArgs(script, payload) {
  const scriptPath = resolveScriptPath(script);
  const wrapperPath = path.posix.join(TOOLBOX_SCRIPT_MOUNT_ROOT.replace(/\\/g, "/"), "Invoke-ToolboxScript.ps1");
  const runtimeScript = { ...script, scriptPath };
  let args;
  let artifacts = {};

  switch (script.id) {
    case "m365-compromised-account-remediation":
      ({ args, artifacts } = buildCompromisedAccountArgs(runtimeScript, payload));
      break;
    case "m365-check-mfa-status":
      ({ args, artifacts } = buildMfaStatusArgs(runtimeScript, payload));
      break;
    case "m365-usage-report":
      ({ args, artifacts } = buildUsageReportArgs(runtimeScript, payload));
      break;
    default:
      if (script.runner === "generic-html") {
        ({ args, artifacts } = buildGenericHtmlArgs(runtimeScript, payload));
        break;
      }
      throw new Error(`No runner is defined for script '${script.id}'.`);
  }

  return {
    scriptPath,
    artifacts,
    args: [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      wrapperPath,
      "-ScriptPath",
      scriptPath,
      ...args
    ]
  };
}

export function listScripts() {
  return scripts;
}

export function getScript(scriptId) {
  return resolveScript(scriptId);
}

export function getRuns() {
  return Array.from(runStore.values()).sort((a, b) => new Date(b.startedAt) - new Date(a.startedAt));
}

export function getRun(runId) {
  return runStore.get(runId) || null;
}

export function startRun(scriptId, payload = {}) {
  ensureOutputDir();

  const script = resolveScript(scriptId);
  const { scriptPath, args, artifacts } = buildArgs(script, payload);
  const runId = uuidv4();
  const run = {
    id: runId,
    scriptId,
    scriptName: script.name,
    status: "running",
    startedAt: new Date().toISOString(),
    finishedAt: null,
    payload,
    command: ["pwsh", ...args].join(" "),
    scriptPath,
    artifacts,
    stdout: "",
    stderr: "",
    exitCode: null
  };

  runStore.set(runId, run);

  const child = spawn("pwsh", args, {
    cwd: OUTPUT_DIR,
    env: process.env
  });

  child.stdout.on("data", (chunk) => {
    run.stdout += chunk.toString();
  });

  child.stderr.on("data", (chunk) => {
    run.stderr += chunk.toString();
  });

  child.on("error", (error) => {
    run.status = "failed";
    run.finishedAt = new Date().toISOString();
    run.stderr += `\n${error.message}`;
  });

  child.on("close", (code) => {
    updateArtifactsFromStdout(run);
    run.exitCode = code;
    run.finishedAt = new Date().toISOString();
    run.status = code === 0 ? "completed" : "failed";
  });

  return run;
}

export function getRunHtml(runId) {
  const run = getRun(runId);
  if (!run) {
    return null;
  }

  const htmlPath = run.artifacts?.htmlPath;
  if (!htmlPath || !fs.existsSync(htmlPath)) {
    return null;
  }

  return {
    path: htmlPath,
    content: fs.readFileSync(htmlPath, "utf8")
  };
}
