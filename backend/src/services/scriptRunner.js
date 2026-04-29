import fs from "node:fs";
import path from "node:path";
import archiver from "archiver";
import { v4 as uuidv4 } from "uuid";
import { OUTPUT_DIR } from "../config/runtime.js";
import { scripts } from "../data/scripts.js";
import { issueArtifactToken } from "./artifactTokens.js";
import {
  addArtifact,
  appendLog,
  createRun,
  deleteRun,
  getRun as getStoredRun,
  getRunArtifactRecord,
  listRuns as listStoredRuns,
  setApproval,
  updateRun
} from "./runStore.js";
import { createError, validatePayload } from "./validation.js";
import { ensureRuntimePaths } from "./runRepository.js";
import { enqueueRun, getQueueMetrics, scriptRunQueue } from "./queue.js";

const scriptsById = new Map(scripts.map((script) => [script.id, script]));

function nowIso() {
  return new Date().toISOString();
}

function isPathInsideDirectory(candidatePath, directoryPath) {
  const relativePath = path.relative(path.resolve(directoryPath), path.resolve(candidatePath));
  return relativePath === "" || (!relativePath.startsWith("..") && !path.isAbsolute(relativePath));
}

function assertArtifactPathAllowed(artifactPath) {
  if (!isPathInsideDirectory(artifactPath, OUTPUT_DIR)) {
    throw createError("Artifact path is outside the allowed output directory.", 403);
  }
}

function resolveScript(scriptId) {
  const script = scriptsById.get(scriptId);
  if (!script) {
    throw createError(`Unknown script '${scriptId}'.`, 404);
  }

  return script;
}

function inferArtifactsFromOutput(run) {
  const discovered = [];
  const seen = new Set();
  const output = [run?.stdout, run?.stderr].filter(Boolean).join("\n");

  for (const match of output.matchAll(/exported to:\s*([^\r\n]+)/gi)) {
    const artifactPath = match[1]?.trim();
    if (!artifactPath || seen.has(artifactPath) || !fs.existsSync(artifactPath)) {
      continue;
    }

    seen.add(artifactPath);
    const stat = fs.statSync(artifactPath);
    discovered.push({
      id: path.basename(artifactPath),
      name: path.basename(artifactPath),
      path: artifactPath,
      type: path.extname(artifactPath).slice(1).toLowerCase() || "file",
      size: stat.size,
      createdAt: stat.mtime.toISOString()
    });
  }

  return discovered;
}

async function syncRunArtifacts(run) {
  const filesByPath = new Map();
  const knownFiles = Array.isArray(run?.artifacts?.files) ? run.artifacts.files : [];

  for (const file of knownFiles) {
    if (file?.path) {
      filesByPath.set(path.resolve(file.path), file);
    }
  }

  const basePath = run?.artifacts?.basePath || run?.result?.artifactBasePath || null;
  if (basePath) {
    const dir = path.dirname(basePath);
    const prefix = path.basename(basePath);
    if (fs.existsSync(dir)) {
      for (const name of fs.readdirSync(dir).filter((entry) => entry.startsWith(prefix))) {
        const fullPath = path.join(dir, name);
        const stat = fs.statSync(fullPath);
        filesByPath.set(path.resolve(fullPath), {
          id: name,
          name,
          path: fullPath,
          type: path.extname(name).slice(1).toLowerCase() || "file",
          size: stat.size,
          createdAt: stat.mtime.toISOString()
        });
      }
    }
  }

  for (const file of inferArtifactsFromOutput(run)) {
    filesByPath.set(path.resolve(file.path), file);
  }

  const files = [...filesByPath.values()].sort((left, right) => left.name.localeCompare(right.name));
  for (const file of files) {
    await addArtifact(run.id, file);
  }

  const htmlFile = files.find((file) => file.type === "html");
  run.artifacts = {
    ...run.artifacts,
    basePath,
    htmlPath: htmlFile?.path || run.artifacts?.htmlPath || null,
    files
  };
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
  await syncRunArtifacts(run);
  const queueMetrics = await getQueueMetrics().catch(() => ({ waiting: 0 }));
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
    queueSize: queueMetrics.waiting || 0,
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

export async function getRuns(filters = {}) {
  const page = await listStoredRuns(filters);
  return {
    ...page,
    items: await Promise.all(page.items.map((run) => cloneForResponse(run)))
  };
}

export async function getRun(runId) {
  const run = await getStoredRun(runId);
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
  const run = await createRun({
    id: uuidv4(),
    scriptId,
    scriptName: script.name,
    mode: script.mode,
    status: "queued",
    requestedAt,
    queuedAt: requestedAt,
    lastActivityAt: requestedAt,
    parameters: validatedPayload,
    currentStep: "Queued for worker execution",
    result: {
      events: [
        {
          type: "progress",
          timestamp: requestedAt,
          message: "Queued for worker execution."
        }
      ],
      metrics: {}
    },
    approval: {
      status: script.approvalRequired ? "approved" : "not_required",
      requestedBy: options.requestedBy || null,
      approvedBy: script.approvalRequired ? options.requestedBy || null : null,
      reason: script.approvalRequired ? "Approved from toolbox UI before launch." : null,
      createdAt: requestedAt
    }
  });

  await appendLog(run.id, "stdout", "Queued for worker execution.", {
    id: uuidv4(),
    level: "progress",
    createdAt: requestedAt
  });
  await enqueueRun(run);
  return getRun(run.id);
}

export async function cancelRun(runId) {
  const run = await getStoredRun(runId);
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

    await updateRun(run.id, {
      status: "canceled",
      finishedAt: timestamp,
      cancelRequestedAt: timestamp,
      lastActivityAt: timestamp,
      currentStep: "Canceled before worker launch",
      summary: "Queued run canceled before worker launch."
    });
    await appendLog(run.id, "system", "Queued run canceled before worker launch.", {
      id: uuidv4(),
      level: "warn",
      createdAt: timestamp
    });
    return getRun(run.id);
  }

  await updateRun(run.id, {
    status: "canceling",
    cancelRequestedAt: timestamp,
    lastActivityAt: timestamp,
    currentStep: "Stopping PowerShell script",
    summary: "Cancellation requested from UI."
  });
  await appendLog(run.id, "system", "Cancellation requested from UI.", {
    id: uuidv4(),
    level: "warn",
    createdAt: timestamp
  });
  return getRun(run.id);
}

export async function getRunArtifacts(runId) {
  const run = await getStoredRun(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  const response = await cloneForResponse(run);
  return response.artifacts.files;
}

export async function getRunArtifact(runId, artifactId) {
  const run = await getStoredRun(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  await syncRunArtifacts(run);
  const file =
    run.artifacts.files.find((entry) => entry.id === artifactId || entry.name === artifactId) ||
    (await getRunArtifactRecord(runId, artifactId));

  if (!file || !fs.existsSync(file.path)) {
    throw createError("Artifact not found for this run.", 404);
  }

  assertArtifactPathAllowed(file.path);

  return file;
}

export async function getRunHtml(runId) {
  const run = await getStoredRun(runId);
  if (!run) {
    return null;
  }

  await syncRunArtifacts(run);
  const htmlArtifact = run.artifacts.files.find((file) => file.type === "html");
  if (!htmlArtifact || !fs.existsSync(htmlArtifact.path)) {
    return null;
  }

  assertArtifactPathAllowed(htmlArtifact.path);

  return {
    path: htmlArtifact.path,
    content: fs.readFileSync(htmlArtifact.path, "utf8")
  };
}

export async function buildArtifactArchive(runId, outputStream) {
  const run = await getStoredRun(runId);
  if (!run) {
    throw createError("Run not found.", 404);
  }

  const artifacts = await syncRunArtifacts(run);
  if (!artifacts.length) {
    throw createError("No artifacts are available for this run.", 404);
  }

  for (const artifact of artifacts) {
    if (artifact.path && fs.existsSync(artifact.path)) {
      assertArtifactPathAllowed(artifact.path);
    }
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

export async function removeRun(runId) {
  await deleteRun(runId);
}

export async function setRunApproval(runId, approvalData) {
  await setApproval(runId, approvalData);
  return getRun(runId);
}
