import path from "node:path";
import assert from "node:assert/strict";
import test from "node:test";
import { prisma } from "./db.js";
import {
  addArtifact,
  appendLog,
  createRun,
  deleteRun,
  getRun,
  listRuns,
  redactSensitiveParameters,
  updateRun
} from "./runStore.js";

function createFakePrisma() {
  const runs = new Map();
  const logs = [];
  const artifacts = [];
  const approvals = new Map();

  function clone(value) {
    return structuredClone(value);
  }

  function hydrateRun(run) {
    if (!run) {
      return null;
    }

    const hydrated = clone(run);
    hydrated.logs = logs.filter((entry) => entry.runId === run.id).map((entry) => clone(entry));
    hydrated.artifacts = artifacts.filter((entry) => entry.runId === run.id).map((entry) => clone(entry));
    hydrated.approval = approvals.has(run.id) ? clone(approvals.get(run.id)) : null;
    return hydrated;
  }

  function matchesWhere(run, where = {}) {
    if (!run) {
      return false;
    }

    if (where.status && run.status !== where.status) {
      return false;
    }
    if (where.scriptId && run.scriptId !== where.scriptId) {
      return false;
    }
    if (where.requestedBy && run.requestedBy !== where.requestedBy) {
      return false;
    }

    if (where.OR?.length) {
      const matched = where.OR.some((condition) => {
        if (condition.tenantId) {
          return run.tenantId === condition.tenantId;
        }
        if (condition.tenantHint?.contains) {
          const haystack = String(run.tenantHint || "").toLowerCase();
          return haystack.includes(String(condition.tenantHint.contains).toLowerCase());
        }
        return false;
      });

      if (!matched) {
        return false;
      }
    }

    if (where.createdAt?.gte && new Date(run.createdAt).getTime() < new Date(where.createdAt.gte).getTime()) {
      return false;
    }
    if (where.createdAt?.lte && new Date(run.createdAt).getTime() > new Date(where.createdAt.lte).getTime()) {
      return false;
    }

    return true;
  }

  return {
    state: { runs, logs, artifacts, approvals },
    run: {
      async create({ data }) {
        const record = {
          ...clone(data),
          createdAt: data.createdAt || new Date(),
          updatedAt: data.updatedAt || data.createdAt || new Date()
        };
        runs.set(record.id, record);
        if (data.approval?.create) {
          approvals.set(record.id, {
            ...clone(data.approval.create),
            runId: record.id,
            updatedAt: data.approval.create.createdAt || new Date()
          });
        }
        return hydrateRun(record);
      },
      async findUnique({ where }) {
        return hydrateRun(runs.get(where.id) || null);
      },
      async count({ where }) {
        return [...runs.values()].filter((run) => matchesWhere(run, where)).length;
      },
      async findMany({ where, skip = 0, take = 25 }) {
        return [...runs.values()]
          .filter((run) => matchesWhere(run, where))
          .sort((left, right) => new Date(right.createdAt).getTime() - new Date(left.createdAt).getTime())
          .slice(skip, skip + take)
          .map((run) => hydrateRun(run));
      },
      async update({ where, data }) {
        const current = runs.get(where.id);
        if (!current) {
          return null;
        }

        const updated = {
          ...current,
          ...clone(data),
          updatedAt: new Date()
        };
        runs.set(where.id, updated);
        return hydrateRun(updated);
      },
      async delete({ where }) {
        runs.delete(where.id);
        approvals.delete(where.id);

        for (let index = logs.length - 1; index >= 0; index -= 1) {
          if (logs[index].runId === where.id) {
            logs.splice(index, 1);
          }
        }

        for (let index = artifacts.length - 1; index >= 0; index -= 1) {
          if (artifacts[index].runId === where.id) {
            artifacts.splice(index, 1);
          }
        }
      }
    },
    runLog: {
      async create({ data }) {
        logs.push(clone(data));
        return clone(data);
      }
    },
    runArtifact: {
      async upsert({ where, update, create }) {
        const index = artifacts.findIndex(
          (entry) => entry.runId === where.runId_path.runId && entry.path === where.runId_path.path
        );
        if (index >= 0) {
          artifacts[index] = {
            ...artifacts[index],
            ...clone(update)
          };
          return clone(artifacts[index]);
        }

        artifacts.push(clone(create));
        return clone(create);
      }
    },
    approval: {
      async upsert({ where, update, create }) {
        const current = approvals.get(where.runId);
        if (current) {
          const updated = {
            ...current,
            ...clone(update),
            updatedAt: new Date()
          };
          approvals.set(where.runId, updated);
          return clone(updated);
        }

        const created = {
          ...clone(create),
          updatedAt: create.createdAt || new Date()
        };
        approvals.set(where.runId, created);
        return clone(created);
      }
    },
    async $transaction(entries) {
      return Promise.all(entries);
    }
  };
}

function withFakePrisma(fn) {
  return async () => {
    const original = {
      run: prisma.run,
      runLog: prisma.runLog,
      runArtifact: prisma.runArtifact,
      approval: prisma.approval,
      $transaction: prisma.$transaction
    };

    const fake = createFakePrisma();
    prisma.run = fake.run;
    prisma.runLog = fake.runLog;
    prisma.runArtifact = fake.runArtifact;
    prisma.approval = fake.approval;
    prisma.$transaction = fake.$transaction;

    try {
      await fn(fake.state);
    } finally {
      prisma.run = original.run;
      prisma.runLog = original.runLog;
      prisma.runArtifact = original.runArtifact;
      prisma.approval = original.approval;
      prisma.$transaction = original.$transaction;
    }
  };
}

test("createRun redacts sensitive parameters and preserves tenant hint", withFakePrisma(async () => {
  const run = await createRun({
    id: "00000000-0000-0000-0000-000000000001",
    scriptId: "m365-test",
    scriptName: "Test Script",
    parameters: {
      tenantId: "contoso.onmicrosoft.com",
      username: "admin@contoso.com",
      clientSecret: "super-secret",
      nested: {
        refreshToken: "refresh-me"
      }
    }
  });

  assert.equal(run.payload.clientSecret, "[REDACTED]");
  assert.equal(run.payload.nested.refreshToken, "[REDACTED]");
  assert.equal(run.tenantHint, "contoso.onmicrosoft.com");
  assert.equal(run.tenantId, null);
  assert.equal(run.canRerun, false);
}));

test("updateRun persists status changes and timing metadata", withFakePrisma(async () => {
  const id = "00000000-0000-0000-0000-000000000002";
  await createRun({ id, scriptId: "m365-test", scriptName: "Test Script" });

  const updated = await updateRun(id, {
    status: "completed",
    startedAt: "2026-04-28T10:00:00.000Z",
    finishedAt: "2026-04-28T10:02:00.000Z",
    exitCode: 0,
    durationMs: 120000
  });

  assert.equal(updated.status, "completed");
  assert.equal(updated.exitCode, 0);
  assert.equal(updated.durationMs, 120000);
  assert.equal(updated.summary, "Completed successfully.");
}));

test("appendLog stores incremental stdout and stderr logs", withFakePrisma(async () => {
  const id = "00000000-0000-0000-0000-000000000003";
  await createRun({ id, scriptId: "m365-test", scriptName: "Test Script" });

  await appendLog(id, "stdout", "[+] Starting report");
  await appendLog(id, "stderr", "[!] Something failed");
  await appendLog(id, "system", "Queued by the API");

  const run = await getRun(id);
  assert.match(run.stdout, /Starting report/);
  assert.match(run.stderr, /Something failed/);
  assert.equal(run.logs.length, 3);
}));

test("addArtifact records artifact metadata and normalizes file names", withFakePrisma(async () => {
  const id = "00000000-0000-0000-0000-000000000004";
  await createRun({ id, scriptId: "m365-test", scriptName: "Test Script" });

  await addArtifact(id, {
    path: path.join(process.cwd(), "output", "report.html"),
    type: "html",
    size: 42
  });

  const run = await getRun(id);
  assert.equal(run.artifacts.files.length, 1);
  assert.equal(run.artifacts.files[0].type, "html");
  assert.equal(run.artifacts.files[0].name, "report.html");
}));

test("addArtifact replaces non-uuid artifact ids before persistence", withFakePrisma(async (state) => {
  const id = "00000000-0000-0000-0000-000000000015";
  await createRun({ id, scriptId: "m365-test", scriptName: "Test Script" });

  await addArtifact(id, {
    id: "Licensing_arnold-rv_28.04.26-20.11.50.html",
    path: path.join(process.cwd(), "output", "Licensing_arnold-rv_28.04.26-20.11.50.html"),
    type: "html",
    size: 99
  });

  assert.equal(state.artifacts.length, 1);
  assert.match(state.artifacts[0].id, /^[0-9a-f-]{36}$/i);
  assert.equal(state.artifacts[0].filename, "Licensing_arnold-rv_28.04.26-20.11.50.html");
}));

test("listRuns supports filters and pagination", withFakePrisma(async () => {
  await createRun({
    id: "00000000-0000-0000-0000-000000000011",
    scriptId: "script-a",
    scriptName: "Script A",
    status: "completed",
    requestedBy: "alice",
    parameters: { tenantId: "contoso.onmicrosoft.com" },
    requestedAt: "2026-04-28T10:00:00.000Z"
  });
  await createRun({
    id: "00000000-0000-0000-0000-000000000012",
    scriptId: "script-b",
    scriptName: "Script B",
    status: "failed",
    requestedBy: "bob",
    parameters: { tenantId: "fabrikam.onmicrosoft.com" },
    requestedAt: "2026-04-28T11:00:00.000Z"
  });
  await createRun({
    id: "00000000-0000-0000-0000-000000000013",
    scriptId: "script-a",
    scriptName: "Script A",
    status: "completed",
    requestedBy: "alice",
    parameters: { tenantId: "contoso.onmicrosoft.com" },
    requestedAt: "2026-04-28T12:00:00.000Z"
  });

  const filtered = await listRuns({
    status: "completed",
    scriptId: "script-a",
    tenantId: "contoso",
    requestedBy: "alice",
    limit: 1,
    offset: 0
  });

  assert.equal(filtered.total, 2);
  assert.equal(filtered.items.length, 1);
  assert.equal(filtered.items[0].scriptId, "script-a");
  assert.equal(filtered.items[0].status, "completed");
}));

test("deleteRun removes stored run history", withFakePrisma(async () => {
  const id = "00000000-0000-0000-0000-000000000014";
  await createRun({ id, scriptId: "m365-test", scriptName: "Test Script" });
  await appendLog(id, "stdout", "hello");
  await deleteRun(id);
  const run = await getRun(id);
  assert.equal(run, null);
}));

test("redactSensitiveParameters handles nested objects and arrays", () => {
  const result = redactSensitiveParameters({
    clientSecret: "secret-1",
    entries: [
      { displayName: "first", password: "secret-2" },
      { refreshToken: "secret-3" }
    ]
  });

  assert.deepEqual(result, {
    clientSecret: "[REDACTED]",
    entries: [
      { displayName: "first", password: "[REDACTED]" },
      { refreshToken: "[REDACTED]" }
    ]
  });
});
