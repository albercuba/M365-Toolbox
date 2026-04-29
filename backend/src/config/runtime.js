import path from "node:path";

export const OUTPUT_DIR = process.env.OUTPUT_DIR || path.resolve(process.cwd(), "../output");
export const TOOLBOX_SCRIPT_MOUNT_ROOT =
  process.env.TOOLBOX_SCRIPT_MOUNT_ROOT || path.resolve(process.cwd(), "../scripts");
export const TOOLBOX_CATALOG_ROOT =
  process.env.TOOLBOX_CATALOG_ROOT || path.join(TOOLBOX_SCRIPT_MOUNT_ROOT, "catalog");
export const RUN_STATE_DIR = process.env.RUN_STATE_DIR || path.join(OUTPUT_DIR, ".toolbox");
export const RUNS_DIR = path.join(RUN_STATE_DIR, "runs");
export const WORKER_HEARTBEAT_FILE = path.join(RUN_STATE_DIR, "worker-heartbeat.json");

export const REDIS_URL = process.env.REDIS_URL || "redis://127.0.0.1:6379";
export const JOB_QUEUE_NAME = process.env.JOB_QUEUE_NAME || "toolbox-script-runs";
export const DLQ_QUEUE_NAME = process.env.DLQ_QUEUE_NAME || "toolbox-script-runs-dlq";
export const WORKER_HEARTBEAT_STALE_MS = Math.max(
  5_000,
  Number(process.env.WORKER_HEARTBEAT_STALE_MS || 30_000)
);

export const RUN_RETENTION_HOURS = Math.max(1, Number(process.env.RUN_RETENTION_HOURS || 168));
export const DEFAULT_READONLY_RETENTION_HOURS = Math.max(
  1,
  Number(process.env.READONLY_RUN_RETENTION_HOURS || RUN_RETENTION_HOURS)
);
export const DEFAULT_REMEDIATION_RETENTION_HOURS = Math.max(
  DEFAULT_READONLY_RETENTION_HOURS,
  Number(process.env.REMEDIATION_RUN_RETENTION_HOURS || 720)
);
export const RUN_TIMEOUT_MS = Math.max(60_000, Number(process.env.RUN_TIMEOUT_MS || 30 * 60 * 1000));
export const RUN_ATTEMPTS = Math.max(1, Number(process.env.RUN_ATTEMPTS || 2));
export const RUN_BACKOFF_MS = Math.max(1_000, Number(process.env.RUN_BACKOFF_MS || 15_000));
export const ARTIFACT_TOKEN_TTL_SECONDS = Math.max(
  60,
  Number(process.env.ARTIFACT_TOKEN_TTL_SECONDS || 15 * 60)
);
if (process.env.NODE_ENV === "production" && !process.env.ARTIFACT_TOKEN_SECRET) {
  throw new Error("ARTIFACT_TOKEN_SECRET must be set in production.");
}
export const ARTIFACT_TOKEN_SECRET = process.env.ARTIFACT_TOKEN_SECRET || "change-me";
