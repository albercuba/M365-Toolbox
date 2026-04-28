import fs from "node:fs";
import path from "node:path";
import { v4 as uuidv4, validate as validateUuid } from "uuid";
import { prisma } from "./db.js";

const REDACTED = "[REDACTED]";
const SENSITIVE_KEY_PATTERN = /(password|secret|token|private.?key|client.?secret|refresh.?token|app.?secret|api.?key)/i;

function asIso(value) {
  if (!value) {
    return null;
  }
  return value instanceof Date ? value.toISOString() : new Date(value).toISOString();
}

function toBigIntOrNull(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const numeric = typeof value === "bigint" ? value : BigInt(Math.trunc(Number(value)));
  return numeric;
}

export function redactSensitiveParameters(value, parentKey = "") {
  if (Array.isArray(value)) {
    return value.map((entry) => redactSensitiveParameters(entry, parentKey));
  }

  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [
        key,
        SENSITIVE_KEY_PATTERN.test(key)
          ? REDACTED
          : redactSensitiveParameters(entry, key)
      ])
    );
  }

  if (parentKey && SENSITIVE_KEY_PATTERN.test(parentKey) && value !== null && value !== undefined && value !== "") {
    return REDACTED;
  }

  return value;
}

function extractTenantInfo(parameters) {
  const tenantHint = parameters?.tenantId ? String(parameters.tenantId).trim() : null;
  const tenantId = tenantHint && validateUuid(tenantHint) ? tenantHint : null;
  return {
    tenantId,
    tenantHint
  };
}

function containsRedactedValue(value) {
  if (Array.isArray(value)) {
    return value.some((entry) => containsRedactedValue(entry));
  }

  if (value && typeof value === "object") {
    return Object.values(value).some((entry) => containsRedactedValue(entry));
  }

  return value === REDACTED;
}

function normalizeRunRecord(run, { includeLogs = true } = {}) {
  const orderedLogs = [...(run.logs || [])].sort(
    (left, right) => new Date(left.createdAt).getTime() - new Date(right.createdAt).getTime()
  );
  const stdout = orderedLogs
    .filter((entry) => entry.stream === "stdout")
    .map((entry) => entry.message)
    .join("\n");
  const stderr = orderedLogs
    .filter((entry) => entry.stream === "stderr")
    .map((entry) => entry.message)
    .join("\n");

  const artifacts = [...(run.artifacts || [])]
    .sort((left, right) => new Date(left.createdAt).getTime() - new Date(right.createdAt).getTime())
    .map((artifact) => ({
      id: artifact.id,
      name: artifact.filename,
      filename: artifact.filename,
      path: artifact.path,
      type: artifact.type,
      size: artifact.sizeBytes ?? 0,
      createdAt: asIso(artifact.createdAt)
    }));

  const htmlArtifact = artifacts.find((artifact) => artifact.type === "html");
  const payload = run.parametersRedacted ?? run.parameters ?? {};

  return {
    id: run.id,
    scriptId: run.scriptId,
    scriptName: run.scriptName,
    mode: run.mode || "read-only",
    status: run.status,
    tenantId: run.tenantId || null,
    tenantHint: run.tenantHint || null,
    requestedBy: run.requestedBy || null,
    payload,
    parameters: payload,
    parametersRedacted: run.parametersRedacted ?? null,
    canRerun: !containsRedactedValue(payload),
    result: run.result ?? null,
    summary: run.summary ?? null,
    requestedAt: asIso(run.createdAt),
    queuedAt: asIso(run.queuedAt),
    startedAt: asIso(run.startedAt),
    finishedAt: asIso(run.finishedAt),
    updatedAt: asIso(run.updatedAt),
    lastActivityAt: asIso(run.lastActivityAt),
    cancelRequestedAt: asIso(run.cancelRequestedAt),
    currentStep: run.currentStep || null,
    errorSummary: run.errorSummary || null,
    exitCode: run.exitCode,
    durationMs: run.durationMs === null || run.durationMs === undefined ? null : Number(run.durationMs),
    command: run.command || null,
    commandArgs: Array.isArray(run.commandArgs) ? run.commandArgs : [],
    scriptPath: run.scriptPath || null,
    stdout,
    stderr,
    logs: includeLogs
      ? orderedLogs.map((entry) => ({
          id: entry.id,
          timestamp: asIso(entry.createdAt),
          createdAt: asIso(entry.createdAt),
          stream: entry.stream,
          level: entry.level || "info",
          message: entry.message
        }))
      : [],
    events: run.result?.events || [],
    approval: run.approval
      ? {
          id: run.approval.id,
          status: run.approval.status,
          requestedBy: run.approval.requestedBy,
          approvedBy: run.approval.approvedBy,
          reason: run.approval.reason,
          createdAt: asIso(run.approval.createdAt),
          updatedAt: asIso(run.approval.updatedAt)
        }
      : null,
    artifacts: {
      basePath: run.result?.artifactBasePath || null,
      htmlPath: htmlArtifact?.path || null,
      files: artifacts
    }
  };
}

function buildRunWhere(filters = {}) {
  const where = {};

  if (filters.status) {
    where.status = filters.status;
  }

  if (filters.scriptId) {
    where.scriptId = filters.scriptId;
  }

  if (filters.requestedBy) {
    where.requestedBy = filters.requestedBy;
  }

  if (filters.tenantId) {
    where.OR = [
      validateUuid(filters.tenantId)
        ? { tenantId: filters.tenantId }
        : null,
      { tenantHint: { contains: filters.tenantId, mode: "insensitive" } }
    ].filter(Boolean);
  }

  if (filters.dateFrom || filters.dateTo) {
    where.createdAt = {};
    if (filters.dateFrom) {
      where.createdAt.gte = new Date(filters.dateFrom);
    }
    if (filters.dateTo) {
      where.createdAt.lte = new Date(filters.dateTo);
    }
  }

  return where;
}

function summarizeRunStatus(run) {
  if (run.summary) {
    return run.summary;
  }

  if (run.status === "failed") {
    return run.errorSummary || "The script failed.";
  }

  if (run.status === "completed") {
    return "Completed successfully.";
  }

  return run.currentStep || null;
}

function toDbPatch(patch = {}) {
  const next = { ...patch };
  const payload = {};

  if ("status" in next) payload.status = next.status;
  if ("mode" in next) payload.mode = next.mode;
  if ("scriptName" in next) payload.scriptName = next.scriptName;
  if ("requestedBy" in next) payload.requestedBy = next.requestedBy;
  if ("startedAt" in next) payload.startedAt = next.startedAt ? new Date(next.startedAt) : null;
  if ("finishedAt" in next) payload.finishedAt = next.finishedAt ? new Date(next.finishedAt) : null;
  if ("queuedAt" in next) payload.queuedAt = next.queuedAt ? new Date(next.queuedAt) : null;
  if ("lastActivityAt" in next) payload.lastActivityAt = next.lastActivityAt ? new Date(next.lastActivityAt) : null;
  if ("cancelRequestedAt" in next) payload.cancelRequestedAt = next.cancelRequestedAt ? new Date(next.cancelRequestedAt) : null;
  if ("currentStep" in next) payload.currentStep = next.currentStep;
  if ("errorSummary" in next) payload.errorSummary = next.errorSummary;
  if ("exitCode" in next) payload.exitCode = next.exitCode;
  if ("durationMs" in next) payload.durationMs = toBigIntOrNull(next.durationMs);
  if ("command" in next) payload.command = next.command;
  if ("commandArgs" in next) payload.commandArgs = next.commandArgs;
  if ("scriptPath" in next) payload.scriptPath = next.scriptPath;
  if ("result" in next) payload.result = next.result;
  if ("summary" in next) payload.summary = next.summary;
  if ("parametersRedacted" in next) payload.parametersRedacted = next.parametersRedacted;
  if ("parameters" in next) payload.parameters = next.parameters;
  if ("tenantId" in next) payload.tenantId = next.tenantId;
  if ("tenantHint" in next) payload.tenantHint = next.tenantHint;

  return payload;
}

export async function createRun(data) {
  const parametersRedacted = redactSensitiveParameters(data.parameters ?? data.payload ?? {});
  const tenantInfo = extractTenantInfo(parametersRedacted);
  const created = await prisma.run.create({
    data: {
      id: data.id || uuidv4(),
      scriptId: data.scriptId,
      scriptName: data.scriptName || null,
      mode: data.mode || "read-only",
      status: data.status || "queued",
      tenantId: tenantInfo.tenantId,
      tenantHint: tenantInfo.tenantHint,
      requestedBy: data.requestedBy || null,
      parameters: null,
      parametersRedacted,
      result: data.result || {},
      summary: data.summary || null,
      exitCode: data.exitCode ?? null,
      startedAt: data.startedAt ? new Date(data.startedAt) : null,
      finishedAt: data.finishedAt ? new Date(data.finishedAt) : null,
      queuedAt: data.queuedAt ? new Date(data.queuedAt) : data.requestedAt ? new Date(data.requestedAt) : new Date(),
      lastActivityAt: data.lastActivityAt ? new Date(data.lastActivityAt) : data.requestedAt ? new Date(data.requestedAt) : new Date(),
      currentStep: data.currentStep || null,
      errorSummary: data.errorSummary || null,
      durationMs: toBigIntOrNull(data.durationMs),
      command: data.command || null,
      commandArgs: data.commandArgs || [],
      scriptPath: data.scriptPath || null,
      createdAt: data.requestedAt ? new Date(data.requestedAt) : new Date(),
      approval: data.approval
        ? {
            create: {
              id: data.approval.id || uuidv4(),
              status: data.approval.status,
              requestedBy: data.approval.requestedBy || null,
              approvedBy: data.approval.approvedBy || null,
              reason: data.approval.reason || null,
              createdAt: data.approval.createdAt ? new Date(data.approval.createdAt) : new Date()
            }
          }
        : undefined
    },
    include: {
      logs: true,
      artifacts: true,
      approval: true
    }
  });

  return normalizeRunRecord(created);
}

export async function getRun(id) {
  const run = await prisma.run.findUnique({
    where: { id },
    include: {
      logs: true,
      artifacts: true,
      approval: true
    }
  });

  return run ? normalizeRunRecord(run) : null;
}

export async function listRuns(filters = {}) {
  const limit = Math.max(1, Math.min(Number(filters.limit || 25), 100));
  const offset = Math.max(0, Number(filters.offset || 0));
  const where = buildRunWhere(filters);

  const [total, runs] = await prisma.$transaction([
    prisma.run.count({ where }),
    prisma.run.findMany({
      where,
      orderBy: { createdAt: "desc" },
      skip: offset,
      take: limit,
      include: {
        logs: true,
        artifacts: true,
        approval: true
      }
    })
  ]);

  return {
    items: runs.map((run) => normalizeRunRecord(run)),
    total,
    limit,
    offset
  };
}

export async function updateRun(id, patch) {
  const current = await prisma.run.findUnique({ where: { id } });
  if (!current) {
    return null;
  }

  const mergedParameters = patch.parametersRedacted ?? current.parametersRedacted ?? {};
  const tenantInfo = extractTenantInfo(mergedParameters);

  const updated = await prisma.run.update({
    where: { id },
    data: {
      ...toDbPatch(patch),
      tenantId: "tenantId" in patch ? patch.tenantId : tenantInfo.tenantId,
      tenantHint: "tenantHint" in patch ? patch.tenantHint : tenantInfo.tenantHint,
      summary: summarizeRunStatus({
        ...current,
        ...patch
      })
    },
    include: {
      logs: true,
      artifacts: true,
      approval: true
    }
  });

  return normalizeRunRecord(updated);
}

export async function appendLog(runId, stream, message, options = {}) {
  if (!message) {
    return null;
  }

  return prisma.runLog.create({
    data: {
      id: options.id || uuidv4(),
      runId,
      stream,
      level: options.level || null,
      message,
      createdAt: options.createdAt ? new Date(options.createdAt) : new Date()
    }
  });
}

export async function addArtifact(runId, artifact) {
  const artifactPath = path.resolve(artifact.path);
  const sizeBytes =
    artifact.sizeBytes ??
    artifact.size ??
    (fs.existsSync(artifactPath) ? fs.statSync(artifactPath).size : null);
  const artifactId = artifact.id && validateUuid(String(artifact.id)) ? String(artifact.id) : uuidv4();

  return prisma.runArtifact.upsert({
    where: {
      runId_path: {
        runId,
        path: artifactPath
      }
    },
    update: {
      type: artifact.type || "other",
      filename: artifact.filename || artifact.name || path.basename(artifactPath),
      sizeBytes
    },
    create: {
      id: artifactId,
      runId,
      type: artifact.type || "other",
      filename: artifact.filename || artifact.name || path.basename(artifactPath),
      path: artifactPath,
      sizeBytes,
      createdAt: artifact.createdAt ? new Date(artifact.createdAt) : new Date()
    }
  });
}

export async function setApproval(runId, approvalData) {
  return prisma.approval.upsert({
    where: { runId },
    update: {
      status: approvalData.status,
      requestedBy: approvalData.requestedBy || null,
      approvedBy: approvalData.approvedBy || null,
      reason: approvalData.reason || null
    },
    create: {
      id: approvalData.id || uuidv4(),
      runId,
      status: approvalData.status,
      requestedBy: approvalData.requestedBy || null,
      approvedBy: approvalData.approvedBy || null,
      reason: approvalData.reason || null,
      createdAt: approvalData.createdAt ? new Date(approvalData.createdAt) : new Date()
    }
  });
}

export async function deleteRun(id) {
  await prisma.run.delete({
    where: { id }
  });
}

export async function getRunArtifactRecord(runId, artifactId) {
  const run = await prisma.run.findUnique({
    where: { id: runId },
    include: { artifacts: true }
  });

  if (!run) {
    return null;
  }

  const artifact = run.artifacts.find((entry) => entry.id === artifactId || entry.filename === artifactId);
  if (!artifact) {
    return null;
  }

  return {
    id: artifact.id,
    name: artifact.filename,
    filename: artifact.filename,
    path: artifact.path,
    type: artifact.type,
    size: artifact.sizeBytes ?? 0,
    createdAt: asIso(artifact.createdAt)
  };
}
