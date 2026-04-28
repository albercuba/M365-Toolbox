import fs from "node:fs";
import path from "node:path";
import {
  OUTPUT_DIR,
  RUNS_DIR,
  RUN_STATE_DIR,
  WORKER_HEARTBEAT_FILE
} from "../config/runtime.js";

export function ensureRuntimePaths() {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.mkdirSync(RUN_STATE_DIR, { recursive: true });
}

function ensureLegacyRunPaths() {
  ensureRuntimePaths();
  fs.mkdirSync(RUNS_DIR, { recursive: true });
}

function getRunFile(runId) {
  return path.join(RUNS_DIR, `${runId}.json`);
}

export function loadLegacyRun(runId) {
  ensureLegacyRunPaths();
  const runPath = getRunFile(runId);
  if (!fs.existsSync(runPath)) {
    return null;
  }

  return JSON.parse(fs.readFileSync(runPath, "utf8"));
}

export function listLegacyRuns() {
  ensureLegacyRunPaths();
  return fs
    .readdirSync(RUNS_DIR, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".json"))
    .map((entry) => loadLegacyRun(path.basename(entry.name, ".json")))
    .filter(Boolean)
    .sort(
      (left, right) =>
        new Date(right.requestedAt || right.startedAt || 0).getTime() -
        new Date(left.requestedAt || left.startedAt || 0).getTime()
    );
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
