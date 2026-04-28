import fs from "node:fs";
import path from "node:path";
import {
  DEFAULT_READONLY_RETENTION_HOURS,
  DEFAULT_REMEDIATION_RETENTION_HOURS,
  TOOLBOX_CATALOG_ROOT
} from "../config/runtime.js";

function enrichScript(script) {
  const mode = script.mode || "read-only";
  return {
    ...script,
    mode,
    approvalRequired: Boolean(script.approvalRequired ?? mode === "remediation"),
    artifactRetentionHours:
      script.artifactRetentionHours ||
      (mode === "remediation"
        ? DEFAULT_REMEDIATION_RETENTION_HOURS
        : DEFAULT_READONLY_RETENTION_HOURS)
  };
}

function loadScriptsFromCatalog() {
  if (!fs.existsSync(TOOLBOX_CATALOG_ROOT)) {
    throw new Error(`Script catalog root not found: ${TOOLBOX_CATALOG_ROOT}`);
  }

  return fs
    .readdirSync(TOOLBOX_CATALOG_ROOT, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".json"))
    .map((entry) => {
      const fullPath = path.join(TOOLBOX_CATALOG_ROOT, entry.name);
      const raw = fs.readFileSync(fullPath, "utf8");
      return enrichScript(JSON.parse(raw));
    })
    .sort((left, right) => left.name.localeCompare(right.name));
}

export const scripts = loadScriptsFromCatalog();
