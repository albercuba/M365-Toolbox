import { Queue, QueueEvents } from "bullmq";
import IORedis from "ioredis";
import {
  DLQ_QUEUE_NAME,
  JOB_QUEUE_NAME,
  REDIS_URL,
  RUN_ATTEMPTS,
  RUN_BACKOFF_MS
} from "../config/runtime.js";

export const redisConnection = new IORedis(REDIS_URL, {
  maxRetriesPerRequest: null,
  enableReadyCheck: true
});

export const scriptRunQueue = new Queue(JOB_QUEUE_NAME, { connection: redisConnection });
export const deadLetterQueue = new Queue(DLQ_QUEUE_NAME, { connection: redisConnection });
export const queueEvents = new QueueEvents(JOB_QUEUE_NAME, { connection: redisConnection.duplicate() });

export const runQueueJobOptions = {
  attempts: RUN_ATTEMPTS,
  removeOnComplete: 1000,
  removeOnFail: false,
  backoff: {
    type: "exponential",
    delay: RUN_BACKOFF_MS
  }
};

export async function enqueueRun(run) {
  await scriptRunQueue.add(
    JOB_QUEUE_NAME,
    {
      runId: run.id,
      scriptId: run.scriptId,
      payload: run.payload
    },
    {
      ...runQueueJobOptions,
      jobId: run.id
    }
  );
}

export async function getQueueMetrics() {
  const counts = await scriptRunQueue.getJobCounts(
    "active",
    "completed",
    "delayed",
    "failed",
    "paused",
    "prioritized",
    "waiting",
    "waiting-children"
  );
  const failedCount = await deadLetterQueue.count();
  return {
    ...counts,
    deadLetter: failedCount
  };
}
