import fs from "node:fs";
import path from "node:path";
import archiver from "archiver";
import { v4 as uuidv4 } from "uuid";
import { OUTPUT_DIR } from "../config/runtime.js";
import { scripts } from "../data/scripts.js";
import { issueArtifactToken } from "./artifactTokens.js";
import { createError, validatePayload } from "./validation.js";
import {
  deleteRun,
  ensureRuntimePaths,
  listRuns,
  loadRun,
  pruneExpiredRuns,
  saveRun
} from "./runRepository.js";
import { enqueueRun, getQueueMetrics, scriptRunQueue } from "./queue.js";

const scriptsById = new Map(scripts.map((script) => [script.id, script]));

function nowIso() {
  return new Date().toISOString();
}

function resolveScript(scriptId) {
  const script = scriptsById.get(scriptId);
  if (!script) {
    throw createError(`Unknown script '${scriptId}'.`, 404);
  }

  return script;
}

function updateArtifactsFromFilesystem(run) {
  if (!run?.artifacts?.basePath) {
    return run?.artifacts?.files || [];
  }

  const dir = path.dirname(run.artifacts.basePath);
  const prefix = path.basename(run.artifacts.basePath);
  if (!fs.existsSync(dir)) {
    return run.artifacts.files || [];
  }

  const files = fs
    .readdirSync(dir)
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
    .sort((left, right) => left.name.localeCompare(right.name));

  run.artifacts.files = files;
  const htmlFile = files.find((file) => file.type === "html");
  run.artifacts.htmlPath = htmlFile?.path || run.artifacts.htmlPath || null;
  return files;
}

async function getQueuePosition(runId) {
  try {
    const waitingJobs = await scriptRunQueue.getJobs(["waiting", "prioritized", "delayed"]);
    const index = waitingJobs.findIndex((job) => job.id === runId);
    return index >= 0 ? index + 1 : null;
  } catch {
    return null;
  }
}

async function cloneForResponse(run) {
  updateArtifactsFromFilesystem(run);
  let queueMetrics = {
    waiting: 0
  };
  try {
    queueMetrics = await getQueueMetrics();
  } catch {
    queueMetrics = { waiting: 0 };
  }
  const queuePosition = run.status === "queued" ? await getQueuePosition(run.id) : null;
  const currentStep =
    run.status === "queued" && queuePosition
      ? `Waiting for execution slot (${queuePosition} of ${queueMetrics.waiting} in queue)`
      : run.status === "canceling"
        ? "Stopping PowerShell script"
        : run.currentStep;

  return {
    ...run,
    currentStep,
    queuePosition,
    queueSize: queueMetrics.waiting,
    artifacts: {
      ...run.artifacts,
      files: (run.artifacts?.files || []).map((file) => ({
        id: file.id,
        name: file.name,
        type: file.type,
        size: file.size,
        createdAt: file.createdAt,
        downloadUrl: `/api/runs/${run.id}/artifacts/${encodeURIComponent(file.id)}?token=${issueArtifactToken({
          runId: run.id,
          artifactId: file.id,
          kind: "download"
        })}`
      })),
      htmlPreviewUrl: run.artifacts?.htmlPath
        ? `/api/runs/${run.id}/html?token=${issueArtifactToken({ runId: run.id, kind: "html" })}`
        : null,
      bundleUrl:
        (run.artifacts?.files || []).length > 0
          ? `/api/runs/${run.id}/package.zip?token=${issueArtifactToken({ runId: run.id, kind: "bundle" })}`
          : null
    }
  };
}

export function listScripts() {
  return scripts;
}

export function getScript(scriptId) {
  return resolveScript(scriptId);
}

export async function getRuns() {
  pruneExpiredRuns(scriptsById);
  const runs = listRuns();
  return Promise.all(runs.map((run) => cloneForResponse(run)));
}

export async function getRun(runId) {
  const run = loadRun(runId);
  if (!run) {
    return null;
  }

  return cloneForResponse(run);
}

export async function startRun(scriptId, payload = {}, options = {}) {
  ensureRuntimePaths();
  const script = resolveScript(scriptId);
  const validatedPayload = validatePayload(script, payload);

  if (script.approvalRequired && !options.approvalConfirmed) {
    throw createError(`'${script.name}' is a remediation workflow. Confirm approval before launch.`, 409);
  }

  const requestedAt = nowIso();
  const run = {
    id: uuidv4(),
    scriptId,
    scriptName: script.name,
    mode: script.mode,
    status: "queued",
    requestedAt,
    queuedAt: requestedAt,
    startedAt: null,
    finishedAt: null,
    updatedAt: requestedAt,
    lastActivityAt: requestedAt,
    payload: validatedPayload,
    command: null,
    commandArgs: [],
    scriptPath: null,
    artifacts: {
      basePath: null,
      htmlPath: null,
      files: []
    },
    stdout: "",
    stderr: "",
    logs: [
      {
        id: uuidv4(),
        timestamp: requestedAt,
        stream: "stdout",
        level: "progress",
        message: "Queued for worker execution."
      }
    ],
    events: [
      {
        type: "progress",
        timestamp: requestedAt,
        message: "Queued for worker execution."
      }
    ],
    currentStep: "Queued for worker execution",
    errorSummary: null,
    exitCode: null,
    durationMs: null,
    cancelRequestedAt: null
  };

  saveRun(run);
  await enqueueRun(run);
  return cloneForResponse(run);
}

export async function cancelRun(runId) {
  const run = loadRun(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  if (["completed", "failed", "canceled", "interrupted"].includes(run.status)) {
    throw createError("This run is already finished.", 409);
  }

  const timestamp = nowIso();
  if (run.status === "queued") {
    const job = await scriptRunQueue.getJob(run.id);
    if (job) {
      await job.remove();
    }
    run.cancelRequestedAt = timestamp;
    run.status = "canceled";
    run.finishedAt = timestamp;
    run.updatedAt = timestamp;
    run.lastActivityAt = timestamp;
    run.logs.push({
      id: uuidv4(),
      timestamp,
      stream: "stderr",
      level: "warn",
      message: "Queued run canceled before worker launch."
    });
    saveRun(run);
    return cloneForResponse(run);
  }

  run.cancelRequestedAt = timestamp;
  run.status = "canceling";
  run.updatedAt = timestamp;
  run.lastActivityAt = timestamp;
  run.logs.push({
    id: uuidv4(),
    timestamp,
    stream: "stderr",
    level: "warn",
    message: "Cancellation requested from UI."
  });
  saveRun(run);
  return cloneForResponse(run);
}

export async function getRunArtifacts(runId) {
  const run = loadRun(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  updateArtifactsFromFilesystem(run);
  saveRun(run);
  return (await cloneForResponse(run)).artifacts.files;
}

export function getRunArtifact(runId, artifactId) {
  const run = loadRun(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  const file = updateArtifactsFromFilesystem(run).find((entry) => entry.id === artifactId);
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
  const run = loadRun(runId);
  if (!run) {
    return null;
  }

  const htmlArtifact = updateArtifactsFromFilesystem(run).find((file) => file.type === "html");
  if (!htmlArtifact || !fs.existsSync(htmlArtifact.path)) {
    return null;
  }

  return {
    path: htmlArtifact.path,
    content: fs.readFileSync(htmlArtifact.path, "utf8")
  };
}

export function buildArtifactArchive(runId, outputStream) {
  const run = loadRun(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  const artifacts = updateArtifactsFromFilesystem(run);
  if (!artifacts.length) {
    throw createError("No artifacts are available for this run.", 404);
  }

  const archive = archiver("zip", { zlib: { level: 9 } });
  archive.on("error", (error) => outputStream.destroy(error));
  archive.pipe(outputStream);

  for (const artifact of artifacts) {
    if (!artifact.path || !fs.existsSync(artifact.path)) {
      continue;
    }
    archive.file(artifact.path, { name: artifact.name });
  }

  archive.finalize();
}

export async function getQueueStatus() {
  return getQueueMetrics();
}
