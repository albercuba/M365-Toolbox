import { spawn } from "node:child_process";
import path from "node:path";
import fs from "node:fs";
import { v4 as uuidv4 } from "uuid";
import { scripts } from "../data/scripts.js";

const OUTPUT_DIR = process.env.OUTPUT_DIR || path.resolve(process.cwd(), "../output");
const SCRIPT_MOUNT_ROOT = process.env.SCRIPT_MOUNT_ROOT || "C:/VSCode/Powershell";
const runStore = new Map();

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

function buildArgs(script, payload) {
  const scriptPath = path.posix.join(SCRIPT_MOUNT_ROOT.replace(/\\/g, "/"), script.scriptRelativePath.replace(/\\/g, "/"));
  const args = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptPath, "-OutputPath", OUTPUT_DIR];

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

  if (payload.csvPath) {
    args.push("-CsvPath", String(payload.csvPath));
  }

  if (payload.tenantId) {
    args.push("-TenantId", String(payload.tenantId));
  }

  if (payload.clientId) {
    args.push("-ClientId", String(payload.clientId));
  }

  if (payload.certificateThumbprint) {
    args.push("-CertificateThumbprint", String(payload.certificateThumbprint));
  }

  if (payload.installMissingModules) {
    args.push("-InstallMissingModules");
  }

  if (payload.includeGeneratedPasswordsInResults) {
    args.push("-IncludeGeneratedPasswordsInResults");
  }

  if (payload.whatIf) {
    args.push("-WhatIf");
  }

  return { scriptPath, args };
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
  const { scriptPath, args } = buildArgs(script, payload);
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
    run.exitCode = code;
    run.finishedAt = new Date().toISOString();
    run.status = code === 0 ? "completed" : "failed";
  });

  return run;
}

