import test from "node:test";
import assert from "node:assert/strict";
import { scripts } from "./scripts.js";

test("script catalog loads JSON metadata from disk", () => {
  assert.ok(Array.isArray(scripts));
  assert.ok(scripts.length >= 50);
});

test("script ids are unique and each script has a relative path", () => {
  const seenIds = new Set();

  for (const script of scripts) {
    assert.ok(script.id);
    assert.ok(script.name);
    assert.ok(script.scriptRelativePath);
    assert.equal(seenIds.has(script.id), false, `Duplicate script id: ${script.id}`);
    seenIds.add(script.id);
  }
});

test("catalog enrichment adds execution mode and retention", () => {
  for (const script of scripts) {
    assert.ok(["read-only", "remediation"].includes(script.mode));
    assert.equal(typeof script.approvalRequired, "boolean");
    assert.equal(typeof script.artifactRetentionHours, "number");
    assert.ok(script.artifactRetentionHours >= 1);
  }
});
