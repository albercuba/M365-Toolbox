import test from "node:test";
import assert from "node:assert/strict";
import { flushOutputBuffers, parseOutputChunk } from "./processEvents.js";

function createRun() {
  return {
    id: "run-test",
    artifacts: { files: [], htmlPath: null, basePath: null },
    logs: [],
    events: [],
    stdout: "",
    stderr: "",
    currentStep: null,
    errorSummary: null
  };
}

test("structured output lines are promoted into run events", () => {
  const run = createRun();
  parseOutputChunk(
    run,
    "stdout",
    Buffer.from('::toolbox::{"type":"progress","message":"Connecting to Graph"}\n')
  );

  assert.equal(run.events.length, 1);
  assert.equal(run.events[0].type, "progress");
  assert.equal(run.currentStep, "Connecting to Graph");
  assert.equal(run.stdout, "");
});

test("artifact events populate artifact inventory", () => {
  const run = createRun();
  parseOutputChunk(
    run,
    "stdout",
    Buffer.from('::toolbox::{"type":"artifact","path":"/app/output/report.html","kind":"html","size":42}\n')
  );

  assert.equal(run.artifacts.files.length, 1);
  assert.equal(run.artifacts.files[0].name, "report.html");
  assert.equal(run.artifacts.htmlPath, "/app/output/report.html");
});

test("plain log output still lands in stdout and structured logs", () => {
  const run = createRun();
  parseOutputChunk(run, "stdout", Buffer.from("[+] Starting report"));
  flushOutputBuffers(run);

  assert.match(run.stdout, /\[\+\] Starting report/);
  assert.equal(run.logs.length, 1);
  assert.equal(run.logs[0].level, "progress");
});
