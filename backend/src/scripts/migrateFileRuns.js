import { pathToFileURL } from "node:url";
import { v4 as uuidv4 } from "uuid";
import { ensureDatabaseReady } from "../services/db.js";
import {
  appendLog,
  createRun,
  getRun,
  addArtifact,
  setApproval,
  updateRun
} from "../services/runStore.js";
import { listLegacyRuns } from "../services/runRepository.js";

const defaultDependencies = {
  appendLog,
  createRun,
  getRun,
  addArtifact,
  setApproval,
  updateRun
};

function approvalStatusForRun(run) {
  if (run.approval?.status) {
    return run.approval.status;
  }
  if (run.mode === "remediation") {
    return "approved";
  }
  return "not_required";
}

export async function importLegacyRun(run, dependencies = defaultDependencies) {
  const {
    appendLog: appendLogEntry,
    createRun: createRunEntry,
    getRun: getStoredRun,
    addArtifact: addArtifactEntry,
    setApproval: setApprovalEntry,
    updateRun: updateStoredRun
  } = dependencies;

  const existing = await getStoredRun(run.id);
  if (existing) {
    return "skipped";
  }

  await createRunEntry({
    id: run.id,
    scriptId: run.scriptId,
    scriptName: run.scriptName,
    mode: run.mode,
    status: run.status,
    requestedAt: run.requestedAt,
    queuedAt: run.queuedAt,
    startedAt: run.startedAt,
    finishedAt: run.finishedAt,
    lastActivityAt: run.lastActivityAt || run.updatedAt,
    currentStep: run.currentStep,
    errorSummary: run.errorSummary,
    exitCode: run.exitCode,
    durationMs: run.durationMs,
    command: run.command,
    commandArgs: run.commandArgs,
    scriptPath: run.scriptPath,
    parameters: run.payload || run.parameters || {},
    result: {
      events: run.events || [],
      metrics: {},
      artifactBasePath: run.artifacts?.basePath || null
    },
    approval: {
      id: run.approval?.id || uuidv4(),
      status: approvalStatusForRun(run),
      requestedBy: run.approval?.requestedBy || null,
      approvedBy: run.approval?.approvedBy || null,
      reason: run.approval?.reason || null,
      createdAt: run.requestedAt || new Date().toISOString()
    }
  });

  for (const entry of run.logs || []) {
    await appendLogEntry(run.id, entry.stream || "system", entry.message, {
      id: entry.id || uuidv4(),
      level: entry.level,
      createdAt: entry.timestamp || entry.createdAt || run.requestedAt || new Date().toISOString()
    });
  }

  for (const artifact of run.artifacts?.files || []) {
    await addArtifactEntry(run.id, artifact);
  }

  if (run.approval) {
    await setApprovalEntry(run.id, {
      id: run.approval.id || uuidv4(),
      status: run.approval.status || approvalStatusForRun(run),
      requestedBy: run.approval.requestedBy || null,
      approvedBy: run.approval.approvedBy || null,
      reason: run.approval.reason || null,
      createdAt: run.approval.createdAt || run.requestedAt || new Date().toISOString()
    });
  }

  await updateStoredRun(run.id, {
    status: run.status,
    startedAt: run.startedAt,
    finishedAt: run.finishedAt,
    lastActivityAt: run.lastActivityAt || run.updatedAt,
    currentStep: run.currentStep,
    errorSummary: run.errorSummary,
    exitCode: run.exitCode,
    durationMs: run.durationMs,
    cancelRequestedAt: run.cancelRequestedAt,
    summary: run.summary
  });

  return "imported";
}

export async function migrateLegacyRuns(legacyRuns = listLegacyRuns(), dependencies = defaultDependencies) {
  let imported = 0;
  let skipped = 0;
  let failed = 0;

  for (const run of legacyRuns) {
    try {
      const result = await importLegacyRun(run, dependencies);
      if (result === "imported") {
        imported += 1;
      } else {
        skipped += 1;
      }
    } catch (error) {
      failed += 1;
      console.error(`Failed to import run ${run.id}: ${error.message}`);
    }
  }

  return { imported, skipped, failed };
}

async function main() {
  await ensureDatabaseReady();
  const summary = await migrateLegacyRuns();
  console.log(
    `Legacy file run migration complete. Imported: ${summary.imported}, skipped: ${summary.skipped}, failed: ${summary.failed}`
  );
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}
