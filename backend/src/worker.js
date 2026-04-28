import { spawn } from "node:child_process";
import { Worker } from "bullmq";
import IORedis from "ioredis";
import { v4 as uuidv4 } from "uuid";
import {
  JOB_QUEUE_NAME,
  OUTPUT_DIR,
  REDIS_URL,
  RUN_ATTEMPTS,
  RUN_TIMEOUT_MS
} from "./config/runtime.js";
import { ensureDatabaseReady } from "./services/db.js";
import { scripts } from "./data/scripts.js";
import {
  addLogEntry,
  flushOutputBuffers,
  markRunActivity,
  nowIso,
  parseOutputChunk
} from "./services/processEvents.js";
import {
  addArtifact,
  appendLog,
  getRun as getStoredRun,
  updateRun
} from "./services/runStore.js";
import { saveWorkerHeartbeat } from "./services/runRepository.js";
import { deadLetterQueue } from "./services/queue.js";
import { buildExecutionPlan } from "./services/scriptExecution.js";

const scriptsById = new Map(scripts.map((script) => [script.id, script]));
const workerRedisConnection = new IORedis(REDIS_URL, {
  maxRetriesPerRequest: null,
  enableReadyCheck: true
});

async function syncRunAggregate(run, state) {
  for (let index = state.persistedLogCount; index < run.logs.length; index += 1) {
    const entry = run.logs[index];
    await appendLog(run.id, entry.stream, entry.message, {
      id: entry.id || uuidv4(),
      level: entry.level,
      createdAt: entry.timestamp
    });
  }
  state.persistedLogCount = run.logs.length;

  const files = run.artifacts?.files || [];
  for (let index = state.persistedArtifactCount; index < files.length; index += 1) {
    const artifact = files[index];
    await addArtifact(run.id, artifact);
  }
  state.persistedArtifactCount = files.length;

  await updateRun(run.id, {
    status: run.status,
    startedAt: run.startedAt,
    finishedAt: run.finishedAt,
    lastActivityAt: run.lastActivityAt,
    currentStep: run.currentStep,
    errorSummary: run.errorSummary,
    exitCode: run.exitCode,
    durationMs: run.durationMs,
    command: run.command,
    commandArgs: run.commandArgs,
    scriptPath: run.scriptPath,
    cancelRequestedAt: run.cancelRequestedAt,
    result: {
      ...(run.result || {}),
      artifactBasePath: run.artifacts?.basePath || run.result?.artifactBasePath || null
    },
    summary: run.summary
  });
}

async function finalizeRun(run, state, status, exitCode = null) {
  run.status = status;
  run.exitCode = exitCode;
  run.finishedAt = nowIso();
  markRunActivity(run, run.finishedAt);
  run.durationMs = run.startedAt
    ? new Date(run.finishedAt).getTime() - new Date(run.startedAt).getTime()
    : null;
  delete run._stdoutBuffer;
  delete run._stderrBuffer;
  await syncRunAggregate(run, state);
}

async function executeRun(job) {
  const run = await getStoredRun(job.data.runId);
  if (!run) {
    throw new Error(`Run '${job.data.runId}' was not found in PostgreSQL.`);
  }

  const script = scriptsById.get(run.scriptId);
  if (!script) {
    throw new Error(`Script '${run.scriptId}' is not defined in the catalog.`);
  }

  const state = {
    persistedLogCount: Array.isArray(run.logs) ? run.logs.length : 0,
    persistedArtifactCount: Array.isArray(run.artifacts?.files) ? run.artifacts.files.length : 0
  };

  const plan = buildExecutionPlan(script, run.payload || {});
  run.status = "running";
  run.startedAt = nowIso();
  run.updatedAt = run.startedAt;
  run.lastActivityAt = run.startedAt;
  run.currentStep = "Preparing PowerShell environment";
  run.scriptPath = plan.scriptPath;
  run.commandArgs = plan.commandArgs;
  run.command = ["pwsh", ...plan.commandArgs].join(" ");
  run.artifacts = {
    ...run.artifacts,
    ...plan.artifacts,
    files: Array.isArray(run.artifacts?.files) ? run.artifacts.files : []
  };
  addLogEntry(run, "stdout", "[+] Launching PowerShell script");
  await syncRunAggregate(run, state);

  await job.updateProgress({
    phase: "running",
    step: run.currentStep,
    runId: run.id
  });

  return new Promise((resolve, reject) => {
    let settled = false;
    let syncChain = Promise.resolve();

    const queueSync = () => {
      syncChain = syncChain
        .then(() => syncRunAggregate(run, state))
        .catch((error) => {
          console.error(`Failed to persist incremental run state for ${run.id}:`, error);
        });
      return syncChain;
    };

    const settle = async (handler) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeoutTimer);
      clearInterval(cancelWatcher);
      clearInterval(persistHeartbeat);

      await syncChain;
      await handler();
    };

    const child = spawn("pwsh", plan.commandArgs, {
      cwd: OUTPUT_DIR,
      env: process.env
    });

    const timeoutTimer = setTimeout(() => {
      if (child.killed) {
        return;
      }

      run.status = "canceling";
      run.errorSummary = `The script exceeded the ${Math.round(RUN_TIMEOUT_MS / 60_000)} minute timeout.`;
      run.currentStep = "Stopping PowerShell script after timeout";
      queueSync();
      child.kill();
    }, RUN_TIMEOUT_MS);

    const cancelWatcher = setInterval(async () => {
      const latest = await getStoredRun(run.id);
      if (latest?.cancelRequestedAt && !child.killed) {
        run.status = "canceling";
        run.cancelRequestedAt = latest.cancelRequestedAt;
        run.currentStep = "Stopping PowerShell script";
        queueSync();
        child.kill();
      }
    }, 1000);

    const persistHeartbeat = setInterval(() => {
      saveWorkerHeartbeat({
        pid: process.pid,
        queue: JOB_QUEUE_NAME,
        status: "running",
        activeRunId: run.id
      });
    }, 5000);

    child.stdout.on("data", (chunk) => {
      parseOutputChunk(run, "stdout", chunk);
      queueSync();
    });

    child.stderr.on("data", (chunk) => {
      parseOutputChunk(run, "stderr", chunk);
      queueSync();
    });

    child.on("error", (error) => {
      settle(async () => {
        addLogEntry(run, "stderr", error.message);
        await finalizeRun(run, state, "failed", 1);
        reject(error);
      }).catch(reject);
    });

    child.on("close", (code) => {
      settle(async () => {
        flushOutputBuffers(run);

        if (run.cancelRequestedAt) {
          addLogEntry(run, "stderr", "[!] Run canceled while worker was executing.");
          await finalizeRun(run, state, "canceled", code);
          resolve({ canceled: true });
          return;
        }

        if (code === 0) {
          addLogEntry(run, "stdout", "[+] Script completed successfully");
          await finalizeRun(run, state, "completed", code);
          await job.updateProgress({
            phase: "completed",
            step: "Completed",
            runId: run.id
          });
          resolve({ completed: true });
          return;
        }

        if (!run.errorSummary) {
          run.errorSummary = "The script exited with an error.";
        }
        await finalizeRun(run, state, "failed", code);
        reject(new Error(run.errorSummary));
      }).catch(reject);
    });
  });
}

await ensureDatabaseReady();

const worker = new Worker(
  JOB_QUEUE_NAME,
  executeRun,
  {
    connection: workerRedisConnection,
    concurrency: Math.max(1, Number(process.env.MAX_CONCURRENT_RUNS || 2))
  }
);

worker.on("ready", () => {
  saveWorkerHeartbeat({
    pid: process.pid,
    queue: JOB_QUEUE_NAME,
    status: "ready",
    activeRunId: null
  });
  console.log("M365 Toolbox worker is ready.");
});

worker.on("completed", () => {
  saveWorkerHeartbeat({
    pid: process.pid,
    queue: JOB_QUEUE_NAME,
    status: "idle",
    activeRunId: null
  });
});

worker.on("failed", async (job, error) => {
  if (!job) {
    return;
  }

  const run = await getStoredRun(job.data.runId);
  if (!run) {
    return;
  }

  const hasRetriesLeft = job.attemptsMade < (job.opts.attempts || RUN_ATTEMPTS);
  const timestamp = nowIso();

  if (hasRetriesLeft) {
    await updateRun(run.id, {
      status: "queued",
      queuedAt: timestamp,
      lastActivityAt: timestamp,
      currentStep: `Retry scheduled (${job.attemptsMade} of ${job.opts.attempts || RUN_ATTEMPTS})`,
      summary: `Worker attempt ${job.attemptsMade} failed. BullMQ will retry this run.`
    });
    await appendLog(run.id, "system", `Worker attempt ${job.attemptsMade} failed. BullMQ will retry this run.`, {
      id: uuidv4(),
      level: "warn",
      createdAt: timestamp
    });
    return;
  }

  await updateRun(run.id, {
    status: run.cancelRequestedAt ? "canceled" : "failed",
    finishedAt: timestamp,
    lastActivityAt: timestamp,
    errorSummary: run.errorSummary || error.message,
    summary: run.errorSummary || error.message
  });
  await appendLog(run.id, "stderr", error.message, {
    id: uuidv4(),
    level: "error",
    createdAt: timestamp
  });
  await deadLetterQueue.add(
    "failed-run",
    {
      runId: run.id,
      scriptId: run.scriptId,
      error: error.message,
      attemptsMade: job.attemptsMade
    },
    { removeOnComplete: 100 }
  );
});

worker.on("error", (error) => {
  saveWorkerHeartbeat({
    pid: process.pid,
    queue: JOB_QUEUE_NAME,
    status: "error",
    activeRunId: null,
    error: error.message
  });
  console.error("Worker error:", error);
});

setInterval(() => {
  saveWorkerHeartbeat({
    pid: process.pid,
    queue: JOB_QUEUE_NAME,
    status: "heartbeat",
    activeRunId: null
  });
}, 10_000).unref();
