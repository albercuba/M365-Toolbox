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
import { scripts } from "./data/scripts.js";
import {
  addLogEntry,
  flushOutputBuffers,
  markRunActivity,
  nowIso,
  parseOutputChunk
} from "./services/processEvents.js";
import {
  loadRun,
  saveRun,
  saveWorkerHeartbeat
} from "./services/runRepository.js";
import { deadLetterQueue } from "./services/queue.js";
import { buildExecutionPlan } from "./services/scriptExecution.js";

const scriptsById = new Map(scripts.map((script) => [script.id, script]));
const workerRedisConnection = new IORedis(REDIS_URL, {
  maxRetriesPerRequest: null,
  enableReadyCheck: true
});

function finalizeRun(run, status, exitCode = null) {
  run.status = status;
  run.exitCode = exitCode;
  run.finishedAt = nowIso();
  markRunActivity(run, run.finishedAt);
  run.durationMs = run.startedAt
    ? new Date(run.finishedAt).getTime() - new Date(run.startedAt).getTime()
    : null;
  delete run._stdoutBuffer;
  delete run._stderrBuffer;
  saveRun(run);
}

async function executeRun(job) {
  const run = loadRun(job.data.runId);
  if (!run) {
    throw new Error(`Run '${job.data.runId}' was not found on disk.`);
  }

  const script = scriptsById.get(run.scriptId);
  if (!script) {
    throw new Error(`Script '${run.scriptId}' is not defined in the catalog.`);
  }

  const plan = buildExecutionPlan(script, run.payload);
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
  saveRun(run);

  await job.updateProgress({
    phase: "running",
    step: run.currentStep,
    runId: run.id
  });

  return new Promise((resolve, reject) => {
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
      saveRun(run);
      child.kill();
    }, RUN_TIMEOUT_MS);

    const cancelWatcher = setInterval(() => {
      const latest = loadRun(run.id);
      if (latest?.cancelRequestedAt && !child.killed) {
        run.status = "canceling";
        run.cancelRequestedAt = latest.cancelRequestedAt;
        run.currentStep = "Stopping PowerShell script";
        saveRun(run);
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
      saveRun(run);
    });

    child.stderr.on("data", (chunk) => {
      parseOutputChunk(run, "stderr", chunk);
      saveRun(run);
    });

    child.on("error", (error) => {
      clearTimeout(timeoutTimer);
      clearInterval(cancelWatcher);
      clearInterval(persistHeartbeat);
      addLogEntry(run, "stderr", error.message);
      finalizeRun(run, "failed", 1);
      reject(error);
    });

    child.on("close", async (code) => {
      clearTimeout(timeoutTimer);
      clearInterval(cancelWatcher);
      clearInterval(persistHeartbeat);
      flushOutputBuffers(run);

      if (run.cancelRequestedAt) {
        addLogEntry(run, "stderr", "[!] Run canceled while worker was executing.");
        finalizeRun(run, "canceled", code);
        resolve({ canceled: true });
        return;
      }

      if (code === 0) {
        addLogEntry(run, "stdout", "[+] Script completed successfully");
        finalizeRun(run, "completed", code);
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
      finalizeRun(run, "failed", code);
      reject(new Error(run.errorSummary));
    });
  });
}

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

  const run = loadRun(job.data.runId);
  if (!run) {
    return;
  }

  const hasRetriesLeft = job.attemptsMade < (job.opts.attempts || RUN_ATTEMPTS);
  const timestamp = nowIso();

  if (hasRetriesLeft) {
    run.status = "queued";
    run.queuedAt = timestamp;
    run.updatedAt = timestamp;
    run.lastActivityAt = timestamp;
    run.currentStep = `Retry scheduled (${job.attemptsMade} of ${job.opts.attempts || RUN_ATTEMPTS})`;
    run.logs.push({
      id: uuidv4(),
      timestamp,
      stream: "stderr",
      level: "warn",
      message: `Worker attempt ${job.attemptsMade} failed. BullMQ will retry this run.`
    });
    saveRun(run);
    return;
  }

  run.status = run.cancelRequestedAt ? "canceled" : "failed";
  run.finishedAt = timestamp;
  run.updatedAt = timestamp;
  run.lastActivityAt = timestamp;
  run.errorSummary = run.errorSummary || error.message;
  run.logs.push({
    id: uuidv4(),
    timestamp,
    stream: "stderr",
    level: "error",
    message: error.message
  });
  saveRun(run);
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
