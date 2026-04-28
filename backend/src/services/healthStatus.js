import fs from "node:fs";
import {
  JOB_QUEUE_NAME,
  OUTPUT_DIR,
  REDIS_URL,
  TOOLBOX_SCRIPT_MOUNT_ROOT,
  WORKER_HEARTBEAT_STALE_MS
} from "../config/runtime.js";
import { prisma } from "./db.js";
import { getQueueMetrics, redisConnection } from "./queue.js";
import { loadWorkerHeartbeat } from "./runRepository.js";

let lastStatus = null;
let lastStatusAt = 0;

function maskRedisUrl(connectionString) {
  try {
    const parsed = new URL(connectionString);
    if (parsed.password) {
      parsed.password = "***";
    }
    return parsed.toString();
  } catch {
    return connectionString;
  }
}

export async function getSystemStatus() {
  if (lastStatus && Date.now() - lastStatusAt < 5000) {
    return lastStatus;
  }

  const outputWritable = (() => {
    try {
      fs.mkdirSync(OUTPUT_DIR, { recursive: true });
      fs.accessSync(OUTPUT_DIR, fs.constants.W_OK);
      return true;
    } catch {
      return false;
    }
  })();

  const scriptsMounted = fs.existsSync(TOOLBOX_SCRIPT_MOUNT_ROOT);
  let redis = {
    available: false,
    url: maskRedisUrl(REDIS_URL),
    error: null
  };
  let queue = {
    ready: false,
    queueName: JOB_QUEUE_NAME,
    active: 0,
    waiting: 0,
    delayed: 0,
    failed: 0,
    deadLetter: 0,
    error: null
  };
  let worker = {
    available: false,
    fresh: false,
    lastHeartbeatAt: null,
    activeRunId: null,
    pid: null,
    status: "missing",
    error: null
  };
  let database = {
    available: false,
    error: null
  };
  const heartbeat = loadWorkerHeartbeat();

  try {
    const pong = await redisConnection.ping();
    redis = { available: pong === "PONG", url: maskRedisUrl(REDIS_URL), error: null };
  } catch (error) {
    redis = { available: false, url: maskRedisUrl(REDIS_URL), error: error.message };
  }

  try {
    await prisma.$queryRawUnsafe("SELECT 1");
    database = {
      available: true,
      error: null
    };
  } catch (error) {
    database = {
      available: false,
      error: error.message
    };
  }

  try {
    const metrics = await getQueueMetrics();
    queue = {
      ready: true,
      queueName: JOB_QUEUE_NAME,
      active: metrics.active || 0,
      waiting: (metrics.waiting || 0) + (metrics.prioritized || 0) + (metrics.delayed || 0),
      delayed: metrics.delayed || 0,
      failed: metrics.failed || 0,
      deadLetter: metrics.deadLetter || 0,
      error: null
    };
  } catch (error) {
    queue.error = error.message;
  }

  if (heartbeat?.updatedAt) {
    const lastHeartbeat = new Date(heartbeat.updatedAt).getTime();
    const fresh = Number.isFinite(lastHeartbeat) && Date.now() - lastHeartbeat <= WORKER_HEARTBEAT_STALE_MS;
    worker = {
      available: true,
      fresh,
      lastHeartbeatAt: heartbeat.updatedAt,
      activeRunId: heartbeat.activeRunId || null,
      pid: heartbeat.pid || null,
      status: heartbeat.status || (fresh ? "ready" : "stale"),
      error: heartbeat.error || null
    };
  }

  const executionAvailable = outputWritable && scriptsMounted && redis.available && database.available && worker.fresh;
  const backendStatus = executionAvailable ? "ok" : "degraded";

  lastStatus = {
    checkedAt: new Date().toISOString(),
    backend: { status: backendStatus },
    paths: {
      outputDir: OUTPUT_DIR,
      outputWritable,
      scriptMountRoot: TOOLBOX_SCRIPT_MOUNT_ROOT,
      scriptsMounted
    },
    execution: {
      mode: "queued-worker",
      available: executionAvailable,
      queueName: JOB_QUEUE_NAME
    },
    database,
    redis,
    queue,
    worker
  };
  lastStatusAt = Date.now();
  return lastStatus;
}
