import { spawn } from "node:child_process";
import path from "node:path";
import fs from "node:fs";
import { v4 as uuidv4 } from "uuid";
import { scripts } from "../data/scripts.js";
import { createError, normalizeListValue, validatePayload } from "./validation.js";

const OUTPUT_DIR = process.env.OUTPUT_DIR || path.resolve(process.cwd(), "../output");
const TOOLBOX_SCRIPT_MOUNT_ROOT = process.env.TOOLBOX_SCRIPT_MOUNT_ROOT || path.resolve(process.cwd(), "../scripts");
const RUN_STATE_DIR = process.env.RUN_STATE_DIR || path.join(OUTPUT_DIR, ".toolbox");
const RUN_STATE_FILE = path.join(RUN_STATE_DIR, "runs.json");
const MAX_CONCURRENT_RUNS = Math.max(1, Number(process.env.MAX_CONCURRENT_RUNS || 2));
const RUN_RETENTION_HOURS = Math.max(1, Number(process.env.RUN_RETENTION_HOURS || 168));
const runStore = new Map();
const childStore = new Map();
const queuedRunIds = [];

function nowIso() {
  return new Date().toISOString();
}

function markRunActivity(run, timestamp = nowIso()) {
  run.updatedAt = timestamp;
  run.lastActivityAt = timestamp;
}

function ensureRuntimePaths() {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.mkdirSync(RUN_STATE_DIR, { recursive: true });
}

function persistRuns() {
  ensureRuntimePaths();
  fs.writeFileSync(
    RUN_STATE_FILE,
    JSON.stringify(Array.from(runStore.values()), null, 2),
    "utf8"
  );
}

function addLogEntry(run, stream, message) {
  if (!message) {
    return;
  }

  const clean = message.replace(/\r/g, "").trim();
  if (!clean) {
    return;
  }

  let level = stream === "stderr" ? "error" : "info";
  if (/^\[\!\]/.test(clean) || /error|failed/i.test(clean)) {
    level = "error";
  } else if (/^\[\*\]/.test(clean) || /warn/i.test(clean)) {
    level = "warn";
  } else if (/^\[\+\]/.test(clean)) {
    level = "progress";
  }

  const entry = {
    id: uuidv4(),
    timestamp: nowIso(),
    stream,
    level,
    message: clean
  };

  run.logs.push(entry);
  markRunActivity(run, entry.timestamp);

  if (level === "progress" || (stream === "stdout" && clean.length > 8)) {
    run.currentStep = clean.replace(/^\[[^\]]+\]\s*/, "");
  }

  if (level === "error") {
    run.errorSummary = clean;
  }
}

function appendChunk(run, stream, chunk) {
  const text = chunk.toString();
  if (stream === "stdout") {
    run.stdout += text;
  } else {
    run.stderr += text;
  }

  text
    .split(/\r?\n/)
    .filter(Boolean)
    .forEach((line) => addLogEntry(run, stream, line));
}

function resolveScript(scriptId) {
  const script = scripts.find((entry) => entry.id === scriptId);
  if (!script) {
    throw createError(`Unknown script '${scriptId}'.`, 404);
  }

  return script;
}

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

function buildArgs(script, payload) {
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
    args: [
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

function updateArtifactsFromStdout(run) {
  if (!run?.stdout) {
    return;
  }

  const htmlMatch = run.stdout.match(/\[\+\]\s+HTML dashboard exported to:\s+(.+)/i);
  if (htmlMatch?.[1]) {
    const actualHtmlPath = htmlMatch[1].trim();
    run.artifacts.htmlPath = actualHtmlPath;

    const htmlName = path.basename(actualHtmlPath);
    const htmlDir = path.dirname(actualHtmlPath);
    if (!Array.isArray(run.artifacts.files)) {
      run.artifacts.files = [];
    }

    const existing = run.artifacts.files.find((file) => file.path === actualHtmlPath || file.name === htmlName);
    if (!existing && fs.existsSync(actualHtmlPath)) {
      const stat = fs.statSync(actualHtmlPath);
      run.artifacts.files.push({
        id: htmlName,
        name: htmlName,
        path: actualHtmlPath,
        type: "html",
        size: stat.size,
        createdAt: stat.mtime.toISOString()
      });
    } else if (!run.artifacts.basePath && htmlDir) {
      run.artifacts.basePath = path.join(htmlDir, path.parse(htmlName).name);
    }
  }
}

function refreshArtifacts(run) {
  updateArtifactsFromStdout(run);
  const basePath = run.artifacts?.basePath;
  if (!basePath) {
    return run.artifacts?.files || [];
  }

  const dir = path.dirname(basePath);
  const prefix = path.basename(basePath);
  if (!fs.existsSync(dir)) {
    return [];
  }

  const files = fs.readdirSync(dir)
    .filter((name) => name.startsWith(prefix))
    .map((name) => {
      const fullPath = path.join(dir, name);
      const stat = fs.statSync(fullPath);
      return {
        id: name,
        name,
        path: fullPath,
        type: path.extname(name).slice(1).toLowerCase() || "file",
        size: stat.size,
        createdAt: stat.mtime.toISOString()
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name));

  const existingFiles = Array.isArray(run.artifacts.files) ? run.artifacts.files : [];
  const mergedFiles = [...files];
  for (const file of existingFiles) {
    if (!mergedFiles.some((entry) => entry.path === file.path || entry.name === file.name)) {
      mergedFiles.push(file);
    }
  }

  run.artifacts.files = mergedFiles;
  const htmlFile = mergedFiles.find((file) => file.type === "html");
  run.artifacts.htmlPath = htmlFile?.path || run.artifacts.htmlPath || null;
  return mergedFiles;
}

function finalizeRun(run, status, exitCode = null) {
  refreshArtifacts(run);
  run.status = status;
  run.exitCode = exitCode;
  run.finishedAt = nowIso();
  markRunActivity(run, run.finishedAt);
  run.durationMs = run.startedAt ? new Date(run.finishedAt).getTime() - new Date(run.startedAt).getTime() : null;
  childStore.delete(run.id);
  persistRuns();
  startNextQueuedRun();
}

function activeRunningCount() {
  return Array.from(runStore.values()).filter((run) => run.status === "running" || run.status === "canceling").length;
}

function startNextQueuedRun() {
  if (activeRunningCount() >= MAX_CONCURRENT_RUNS) {
    return;
  }

  const nextRunId = queuedRunIds.shift();
  if (!nextRunId) {
    persistRuns();
    return;
  }

  const run = runStore.get(nextRunId);
  if (!run || run.status !== "queued") {
    startNextQueuedRun();
    return;
  }

  launchRun(run);
}

function launchRun(run) {
  run.status = "running";
  run.startedAt = nowIso();
  markRunActivity(run, run.startedAt);
  run.currentStep = "Preparing PowerShell environment";
  addLogEntry(run, "stdout", "[+] Launching PowerShell script");
  persistRuns();

  const child = spawn("pwsh", run.commandArgs, {
    cwd: OUTPUT_DIR,
    env: process.env
  });
  childStore.set(run.id, child);

  child.stdout.on("data", (chunk) => {
    appendChunk(run, "stdout", chunk);
    persistRuns();
  });

  child.stderr.on("data", (chunk) => {
    appendChunk(run, "stderr", chunk);
    persistRuns();
  });

  child.on("error", (error) => {
    appendChunk(run, "stderr", Buffer.from(error.message));
    finalizeRun(run, "failed", 1);
  });

  child.on("close", (code) => {
    updateArtifactsFromStdout(run);
    if (run.cancelRequestedAt) {
      finalizeRun(run, "canceled", code);
      return;
    }

    if (code === 0) {
      addLogEntry(run, "stdout", "[+] Script completed successfully");
      finalizeRun(run, "completed", code);
      return;
    }

    if (!run.errorSummary) {
      run.errorSummary = "The script exited with an error.";
    }
    finalizeRun(run, "failed", code);
  });
}

function pruneExpiredRuns() {
  const cutoff = Date.now() - RUN_RETENTION_HOURS * 60 * 60 * 1000;
  let changed = false;

  for (const [runId, run] of runStore.entries()) {
    const comparison = run.finishedAt || run.updatedAt || run.requestedAt;
    if (!comparison) {
      continue;
    }

    if (new Date(comparison).getTime() < cutoff) {
      runStore.delete(runId);
      changed = true;
    }
  }

  if (changed) {
    persistRuns();
  }
}

function loadPersistedRuns() {
  ensureRuntimePaths();
  if (!fs.existsSync(RUN_STATE_FILE)) {
    return;
  }

  try {
    const raw = fs.readFileSync(RUN_STATE_FILE, "utf8");
    const savedRuns = JSON.parse(raw);
    for (const savedRun of savedRuns) {
      const run = {
        ...savedRun,
        logs: Array.isArray(savedRun.logs) ? savedRun.logs : [],
        artifacts: savedRun.artifacts || { files: [] }
      };

      if (!run.lastActivityAt) {
        run.lastActivityAt = run.updatedAt || run.finishedAt || run.startedAt || run.requestedAt || nowIso();
      }

      if (run.status === "running" || run.status === "queued" || run.status === "canceling") {
        run.status = "interrupted";
        run.finishedAt = nowIso();
        markRunActivity(run, run.finishedAt);
        run.errorSummary = "The backend restarted before this run completed.";
        run.logs.push({
          id: uuidv4(),
          timestamp: run.finishedAt,
          stream: "stderr",
          level: "error",
          message: "The backend restarted before this run completed."
        });
      }

      refreshArtifacts(run);
      runStore.set(run.id, run);
    }
  } catch (error) {
    console.error("Failed to load persisted runs:", error);
  }

  pruneExpiredRuns();
}

loadPersistedRuns();
setInterval(pruneExpiredRuns, 60 * 60 * 1000).unref();

function getQueuePosition(runId) {
  const queueIndex = queuedRunIds.indexOf(runId);
  return queueIndex >= 0 ? queueIndex + 1 : null;
}

function cloneForResponse(run) {
  refreshArtifacts(run);
  const queuePosition = run.status === "queued" ? getQueuePosition(run.id) : null;
  const queueSize = queuedRunIds.length;
  let currentStep = run.currentStep;

  if (run.status === "queued" && queuePosition) {
    currentStep = `Waiting for execution slot (${queuePosition} of ${queueSize} in queue)`;
  } else if (run.status === "canceling") {
    currentStep = "Stopping PowerShell script";
  }

  return {
    ...run,
    currentStep,
    lastActivityAt: run.lastActivityAt || run.updatedAt || run.requestedAt,
    queuePosition,
    queueSize,
    commandArgs: undefined,
    artifacts: {
      ...run.artifacts,
      files: (run.artifacts?.files || []).map((file) => ({
        id: file.id,
        name: file.name,
        type: file.type,
        size: file.size,
        createdAt: file.createdAt
      }))
    }
  };
}

export function listScripts() {
  return scripts;
}

export function getScript(scriptId) {
  return resolveScript(scriptId);
}

export function getRuns() {
  return Array.from(runStore.values())
    .sort((a, b) => new Date(b.requestedAt || b.startedAt || 0) - new Date(a.requestedAt || a.startedAt || 0))
    .map(cloneForResponse);
}

export function getRun(runId) {
  const run = runStore.get(runId);
  return run ? cloneForResponse(run) : null;
}

export function startRun(scriptId, payload = {}, options = {}) {
  ensureRuntimePaths();
  const script = resolveScript(scriptId);
  const validatedPayload = validatePayload(script, payload);

  if (script.approvalRequired && !options.approvalConfirmed) {
    throw createError(`'${script.name}' is a remediation workflow. Confirm approval before launch.`, 409);
  }

  const { scriptPath, args, artifacts } = buildArgs(script, validatedPayload);
  const run = {
    id: uuidv4(),
    scriptId,
    scriptName: script.name,
    mode: script.mode,
    status: activeRunningCount() >= MAX_CONCURRENT_RUNS ? "queued" : "running",
    requestedAt: nowIso(),
    queuedAt: null,
    startedAt: null,
    finishedAt: null,
    updatedAt: nowIso(),
    lastActivityAt: null,
    payload: validatedPayload,
    command: ["pwsh", ...args].join(" "),
    commandArgs: args,
    scriptPath,
    artifacts,
    stdout: "",
    stderr: "",
    logs: [],
    currentStep: activeRunningCount() >= MAX_CONCURRENT_RUNS ? "Waiting for an execution slot" : "Preparing run",
    errorSummary: null,
    exitCode: null,
    durationMs: null,
    cancelRequestedAt: null
  };

  run.lastActivityAt = run.updatedAt;

  addLogEntry(run, "stdout", run.status === "queued"
    ? `[+] Queued for execution. Concurrency limit is ${MAX_CONCURRENT_RUNS}.`
    : "[+] Run accepted by backend");

  if (run.status === "queued") {
    run.queuedAt = run.requestedAt;
    queuedRunIds.push(run.id);
    run.currentStep = `Waiting for execution slot (${getQueuePosition(run.id)} of ${queuedRunIds.length} in queue)`;
  }

  runStore.set(run.id, run);
  persistRuns();

  if (run.status === "running") {
    launchRun(run);
  }

  return cloneForResponse(run);
}

export function cancelRun(runId) {
  const run = runStore.get(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  if (["completed", "failed", "canceled", "interrupted"].includes(run.status)) {
    throw createError("This run is already finished.", 409);
  }

  if (run.status === "queued") {
    const index = queuedRunIds.indexOf(run.id);
    if (index >= 0) {
      queuedRunIds.splice(index, 1);
    }
    run.cancelRequestedAt = nowIso();
    markRunActivity(run, run.cancelRequestedAt);
    addLogEntry(run, "stderr", "[!] Queued run canceled before launch");
    finalizeRun(run, "canceled", null);
    return cloneForResponse(run);
  }

  const child = childStore.get(run.id);
  if (!child) {
    throw createError("Unable to cancel this run because no active process was found.", 409);
  }

  run.cancelRequestedAt = nowIso();
  run.status = "canceling";
  markRunActivity(run, run.cancelRequestedAt);
  addLogEntry(run, "stderr", "[!] Cancellation requested from UI");
  child.kill();
  persistRuns();
  return cloneForResponse(run);
}

export function getRunArtifacts(runId) {
  const run = runStore.get(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  refreshArtifacts(run);
  persistRuns();
  return run.artifacts.files.map((file) => ({
    id: file.id,
    name: file.name,
    type: file.type,
    size: file.size,
    createdAt: file.createdAt
  }));
}

export function getRunArtifact(runId, artifactId) {
  const run = runStore.get(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  const file = refreshArtifacts(run).find((entry) => entry.id === artifactId);
  if (!file || !fs.existsSync(file.path)) {
    throw createError("Artifact not found for this run.", 404);
  }

  const normalizedOutputDir = path.resolve(OUTPUT_DIR);
  const resolvedArtifactPath = path.resolve(file.path);
  if (!resolvedArtifactPath.startsWith(normalizedOutputDir)) {
    throw createError("Artifact path is outside the allowed output directory.", 403);
  }

  return file;
}

export function getRunHtml(runId) {
  const run = runStore.get(runId);
  if (!run) {
    return null;
  }

  const htmlArtifact = refreshArtifacts(run).find((file) => file.type === "html");
  if (!htmlArtifact || !fs.existsSync(htmlArtifact.path)) {
    return null;
  }

  return {
    path: htmlArtifact.path,
    content: fs.readFileSync(htmlArtifact.path, "utf8")
  };
}
