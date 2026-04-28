import fs from "node:fs";
import path from "node:path";
import {
  OUTPUT_DIR,
  RUNS_DIR,
  RUN_STATE_DIR,
  RUN_RETENTION_HOURS,
  WORKER_HEARTBEAT_FILE
} from "../config/runtime.js";

export function ensureRuntimePaths() {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.mkdirSync(RUN_STATE_DIR, { recursive: true });
  fs.mkdirSync(RUNS_DIR, { recursive: true });
}

function getRunFile(runId) {
  return path.join(RUNS_DIR, `${runId}.json`);
}

export function saveRun(run) {
  ensureRuntimePaths();
  fs.writeFileSync(getRunFile(run.id), JSON.stringify(run, null, 2), "utf8");
}

export function loadRun(runId) {
  ensureRuntimePaths();
  const runPath = getRunFile(runId);
  if (!fs.existsSync(runPath)) {
    return null;
  }

  const parsed = JSON.parse(fs.readFileSync(runPath, "utf8"));
  parsed.logs = Array.isArray(parsed.logs) ? parsed.logs : [];
  parsed.events = Array.isArray(parsed.events) ? parsed.events : [];
  parsed.artifacts = parsed.artifacts || { files: [], htmlPath: null, basePath: null };
  parsed.artifacts.files = Array.isArray(parsed.artifacts.files) ? parsed.artifacts.files : [];
  return parsed;
}

export function listRuns() {
  ensureRuntimePaths();
  return fs
    .readdirSync(RUNS_DIR, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".json"))
    .map((entry) => loadRun(path.basename(entry.name, ".json")))
    .filter(Boolean)
    .sort(
      (left, right) =>
        new Date(right.requestedAt || right.startedAt || 0).getTime() -
        new Date(left.requestedAt || left.startedAt || 0).getTime()
    );
}

export function deleteRun(runId) {
  const runPath = getRunFile(runId);
  if (fs.existsSync(runPath)) {
    fs.unlinkSync(runPath);
  }
}

export function saveWorkerHeartbeat(payload) {
  ensureRuntimePaths();
  fs.writeFileSync(
    WORKER_HEARTBEAT_FILE,
    JSON.stringify({ ...payload, updatedAt: new Date().toISOString() }, null, 2),
    "utf8"
  );
}

export function loadWorkerHeartbeat() {
  ensureRuntimePaths();
  if (!fs.existsSync(WORKER_HEARTBEAT_FILE)) {
    return null;
  }

  return JSON.parse(fs.readFileSync(WORKER_HEARTBEAT_FILE, "utf8"));
}

export function pruneExpiredRuns(scriptsById) {
  const cutoff = Date.now() - RUN_RETENTION_HOURS * 60 * 60 * 1000;
  for (const run of listRuns()) {
    const comparison = run.finishedAt || run.updatedAt || run.requestedAt;
    if (!comparison) {
      continue;
    }

    const script = scriptsById.get(run.scriptId);
    const retentionHours = script?.artifactRetentionHours || RUN_RETENTION_HOURS;
    const scriptCutoff = Date.now() - retentionHours * 60 * 60 * 1000;
    if (new Date(comparison).getTime() >= scriptCutoff || new Date(comparison).getTime() >= cutoff) {
      continue;
    }

    for (const artifact of run.artifacts?.files || []) {
      try {
        if (artifact.path && fs.existsSync(artifact.path)) {
          fs.unlinkSync(artifact.path);
        }
      } catch {
        // Keep pruning resilient.
      }
    }

    deleteRun(run.id);
  }
}
