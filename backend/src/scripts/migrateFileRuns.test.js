import assert from "node:assert/strict";
import test from "node:test";
import { importLegacyRun, migrateLegacyRuns } from "./migrateFileRuns.js";

function createDependencies() {
  const runs = new Map();
  const logs = [];
  const artifacts = [];
  const approvals = [];
  const updates = [];

  return {
    state: { runs, logs, artifacts, approvals, updates },
    deps: {
      async getRun(id) {
        return runs.get(id) || null;
      },
      async createRun(run) {
        runs.set(run.id, { ...run });
      },
      async appendLog(runId, stream, message, options) {
        logs.push({ runId, stream, message, ...options });
      },
      async addArtifact(runId, artifact) {
        artifacts.push({ runId, ...artifact });
      },
      async setApproval(runId, approval) {
        approvals.push({ runId, ...approval });
      },
      async updateRun(id, patch) {
        updates.push({ id, ...patch });
      }
    }
  };
}

function createLegacyRun(id, overrides = {}) {
  return {
    id,
    scriptId: "m365-test",
    scriptName: "Test Script",
    mode: "read-only",
    status: "completed",
    requestedAt: "2026-04-28T10:00:00.000Z",
    queuedAt: "2026-04-28T10:00:01.000Z",
    startedAt: "2026-04-28T10:00:02.000Z",
    finishedAt: "2026-04-28T10:00:05.000Z",
    lastActivityAt: "2026-04-28T10:00:05.000Z",
    currentStep: "Completed",
    exitCode: 0,
    durationMs: 3000,
    payload: {
      tenantId: "contoso.onmicrosoft.com"
    },
    logs: [
      {
        id: `${id}-log-1`,
        stream: "stdout",
        message: "[+] Done",
        createdAt: "2026-04-28T10:00:05.000Z"
      }
    ],
    artifacts: {
      basePath: "/app/output/report",
      files: [
        {
          id: `${id}-artifact-1`,
          path: "/app/output/report.html",
          type: "html",
          name: "report.html"
        }
      ]
    },
    approval: {
      id: `${id}-approval-1`,
      status: "not_required"
    },
    ...overrides
  };
}

test("importLegacyRun imports a legacy file-backed run into the new store", async () => {
  const { state, deps } = createDependencies();
  const run = createLegacyRun("00000000-0000-0000-0000-000000000101");

  const result = await importLegacyRun(run, deps);

  assert.equal(result, "imported");
  assert.equal(state.runs.size, 1);
  assert.equal(state.logs.length, 1);
  assert.equal(state.artifacts.length, 1);
  assert.equal(state.approvals.length, 1);
  assert.equal(state.updates.length, 1);
});

test("importLegacyRun skips duplicates by run id", async () => {
  const { state, deps } = createDependencies();
  const run = createLegacyRun("00000000-0000-0000-0000-000000000102");
  state.runs.set(run.id, { id: run.id });

  const result = await importLegacyRun(run, deps);

  assert.equal(result, "skipped");
  assert.equal(state.logs.length, 0);
  assert.equal(state.artifacts.length, 0);
});

test("migrateLegacyRuns reports imported skipped and failed totals", async () => {
  const { state, deps } = createDependencies();
  const importedRun = createLegacyRun("00000000-0000-0000-0000-000000000103");
  const skippedRun = createLegacyRun("00000000-0000-0000-0000-000000000104");
  const failedRun = createLegacyRun("00000000-0000-0000-0000-000000000105", { id: null });

  state.runs.set(skippedRun.id, { id: skippedRun.id });
  const originalConsoleError = console.error;
  console.error = () => {};

  try {
    const summary = await migrateLegacyRuns([importedRun, skippedRun, failedRun], {
      ...deps,
      async createRun(run) {
        if (!run.id) {
          throw new Error("Missing id");
        }
        state.runs.set(run.id, { ...run });
      }
    });

    assert.deepEqual(summary, {
      imported: 1,
      skipped: 1,
      failed: 1
    });
  } finally {
    console.error = originalConsoleError;
  }
});
