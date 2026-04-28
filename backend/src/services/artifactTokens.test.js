import test from "node:test";
import assert from "node:assert/strict";
import { issueArtifactToken, verifyArtifactToken } from "./artifactTokens.js";

test("artifact tokens verify for the intended run and artifact", () => {
  const token = issueArtifactToken({
    runId: "run-123",
    artifactId: "report.html",
    kind: "download"
  });

  assert.equal(
    verifyArtifactToken(token, {
      runId: "run-123",
      artifactId: "report.html",
      kind: "download"
    }),
    true
  );
});

test("artifact tokens fail verification when scope changes", () => {
  const token = issueArtifactToken({
    runId: "run-123",
    artifactId: "report.html",
    kind: "download"
  });

  assert.equal(
    verifyArtifactToken(token, {
      runId: "run-123",
      artifactId: "report.csv",
      kind: "download"
    }),
    false
  );
  assert.equal(
    verifyArtifactToken(token, {
      runId: "run-999",
      artifactId: "report.html",
      kind: "download"
    }),
    false
  );
});
